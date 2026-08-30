// These are the operations that touch more than one document atomically —
// the same ones that needed real transactions in the Postgres backend
// (receiving a PO, completing a production batch, delivering a sales
// order, transferring stock). Firestore's runTransaction() gives the same
// all-or-nothing guarantee: if a stock check fails partway through, nothing
// commits, so you never end up with half-updated inventory.
import { runTransaction, doc, collection, serverTimestamp } from "firebase/firestore";
import { db } from "./firebase";

// IDs are generated client-side (timestamp + random suffix) rather than the
// old in-memory version's simple incrementing counter — that counter reset
// every page load, which was fine for a single local session but would
// collide across two real, concurrent browser sessions now that everyone
// shares the same Firestore database. setDoc() with a colliding ID would
// silently overwrite someone else's record instead of erroring.
export const uid = (prefix) => `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
export const todayStr = () => new Date().toISOString().slice(0, 10);

export async function receivePurchaseOrder(po) {
  return runTransaction(db, async (tx) => {
    const materialRefs = po.items.map((it) => doc(db, "materials", it.id));
    const materialSnaps = await Promise.all(materialRefs.map((ref) => tx.get(ref)));

    materialSnaps.forEach((snap, i) => {
      if (!snap.exists()) throw new Error(`Material ${po.items[i].id} not found.`);
    });

    materialSnaps.forEach((snap, i) => {
      const current = snap.data().stock || 0;
      tx.update(materialRefs[i], { stock: current + po.items[i].qty });
    });

    po.items.forEach((it) => {
      tx.set(doc(collection(db, "movements")), {
        type: "IN", itemType: "Material", item: it.id, qty: it.qty,
        warehouse: "BR-HQ", date: todayStr(), ref: po.id,
      });
    });

    tx.update(doc(db, "purchaseOrders", po.id), { status: "Received" });

    const amount = po.items.reduce((s, it) => s + it.qty * it.cost, 0);
    const invoiceId = uid("PINV");
    tx.set(doc(db, "purchaseInvoices", invoiceId), {
      po: po.id, supplier: po.supplier, amount, date: todayStr(), status: "Unpaid",
    });
    return invoiceId;
  });
}

export async function completeProductionBatch(batch, product, matById) {
  return runTransaction(db, async (tx) => {
    const materialRefs = product.bom.map((b) => doc(db, "materials", b.id));
    const materialSnaps = await Promise.all(materialRefs.map((ref) => tx.get(ref)));
    const productRef = doc(db, "products", product.id);
    const productSnap = await tx.get(productRef);

    product.bom.forEach((b, i) => {
      const needed = b.qty * batch.qty;
      const have = materialSnaps[i].data().stock || 0;
      if (have < needed) throw new Error(`Not enough ${matById[b.id]?.name || b.id} to complete this batch (need ${needed}, have ${have}).`);
    });

    product.bom.forEach((b, i) => {
      const needed = +(b.qty * batch.qty).toFixed(2);
      const current = materialSnaps[i].data().stock || 0;
      tx.update(materialRefs[i], { stock: +(current - needed).toFixed(2) });
      tx.set(doc(collection(db, "movements")), {
        type: "OUT", itemType: "Material", item: b.id, qty: needed,
        warehouse: "BR-HQ", date: todayStr(), ref: batch.id,
      });
    });

    const stockByBranch = productSnap.data().stockByBranch || {};
    tx.update(productRef, { stockByBranch: { ...stockByBranch, "BR-HQ": (stockByBranch["BR-HQ"] || 0) + batch.qty } });
    tx.set(doc(collection(db, "movements")), {
      type: "IN", itemType: "Product", item: product.id, qty: batch.qty,
      warehouse: "BR-HQ", date: todayStr(), ref: batch.id,
    });

    tx.update(doc(db, "productionOrders", batch.id), { status: "Completed" });
  });
}

export async function transferProductStock(productId, from, to, qty) {
  return runTransaction(db, async (tx) => {
    const productRef = doc(db, "products", productId);
    const snap = await tx.get(productRef);
    if (!snap.exists()) throw new Error("Product not found.");
    const stockByBranch = snap.data().stockByBranch || {};
    if ((stockByBranch[from] || 0) < qty) throw new Error(`Not enough stock at ${from} to transfer.`);

    tx.update(productRef, {
      stockByBranch: {
        ...stockByBranch,
        [from]: stockByBranch[from] - qty,
        [to]: (stockByBranch[to] || 0) + qty,
      },
    });
    tx.set(doc(collection(db, "transfers")), { product: productId, from, to, qty, date: todayStr() });
  });
}

// Sales order Pending -> Delivered: checks stock is actually available
// before deducting — this is the exact bug that was found and fixed in the
// web app's in-memory version (a demo order for more units than were in
// stock silently went negative). Firestore transactions make it structural
// here: the check and the deduction happen in the same atomic operation.
export async function deliverSalesOrder(order, branchName) {
  return runTransaction(db, async (tx) => {
    const productRefs = order.items.map((it) => doc(db, "products", it.id));
    const productSnaps = await Promise.all(productRefs.map((ref) => tx.get(ref)));

    productSnaps.forEach((snap, i) => {
      const item = order.items[i];
      const stockByBranch = snap.data()?.stockByBranch || {};
      const available = stockByBranch[order.branch] || 0;
      if (available < item.qty) {
        throw new Error(`Not enough stock at ${branchName} to deliver this order (need ${item.qty}, have ${available}).`);
      }
    });

    productSnaps.forEach((snap, i) => {
      const item = order.items[i];
      const stockByBranch = snap.data().stockByBranch || {};
      tx.update(productRefs[i], { stockByBranch: { ...stockByBranch, [order.branch]: stockByBranch[order.branch] - item.qty } });
      tx.set(doc(collection(db, "movements")), {
        type: "OUT", itemType: "Product", item: item.id, qty: item.qty,
        warehouse: order.branch, date: todayStr(), ref: order.id,
      });
    });

    tx.update(doc(db, "salesOrders", order.id), { status: "Delivered" });
  });
}

export async function restockReturn(ret, branch) {
  return runTransaction(db, async (tx) => {
    const productRef = doc(db, "products", ret.item);
    const snap = await tx.get(productRef);
    const stockByBranch = snap.data().stockByBranch || {};
    tx.update(productRef, { stockByBranch: { ...stockByBranch, [branch]: (stockByBranch[branch] || 0) + ret.qty } });
    tx.set(doc(collection(db, "movements")), {
      type: "IN", itemType: "Product", item: ret.item, qty: ret.qty,
      warehouse: branch, date: todayStr(), ref: ret.id,
    });
    tx.update(doc(db, "returns", ret.id), { status: "Restocked" });
  });
}

export async function convertPurchaseRequestToOrder(pr, material) {
  return runTransaction(db, async (tx) => {
    const poId = uid("PO");
    tx.set(doc(db, "purchaseOrders", poId), {
      supplier: material.supplier, date: todayStr(), status: "Pending",
      items: [{ id: pr.item, qty: pr.qty, cost: material.cost }],
    });
    tx.update(doc(db, "purchaseRequests", pr.id), { status: "Converted" });
    return poId;
  });
}

export async function convertQuotationToOrder(quotation, client, prodById) {
  return runTransaction(db, async (tx) => {
    const soId = uid("SO");
    tx.set(doc(db, "salesOrders", soId), {
      client: quotation.client, branch: client.branch, date: todayStr(), status: "Pending",
      items: quotation.items.map((it) => ({ id: it.id, qty: it.qty, price: prodById[it.id]?.price ?? 0 })),
    });
    tx.update(doc(db, "quotations", quotation.id), { status: "Converted" });
    return soId;
  });
}

export async function addStoreItemTx({ itemId, warehouse, qty, reason, isMaterial }) {
  return runTransaction(db, async (tx) => {
    if (isMaterial) {
      const ref = doc(db, "materials", itemId);
      const snap = await tx.get(ref);
      tx.update(ref, { stock: (snap.data().stock || 0) + qty });
    } else {
      const ref = doc(db, "products", itemId);
      const snap = await tx.get(ref);
      const stockByBranch = snap.data().stockByBranch || {};
      tx.update(ref, { stockByBranch: { ...stockByBranch, [warehouse]: (stockByBranch[warehouse] || 0) + qty } });
    }
    tx.set(doc(collection(db, "movements")), {
      type: "IN", itemType: isMaterial ? "Material" : "Product", item: itemId, qty,
      warehouse, date: todayStr(), ref: reason || "Manual Add",
    });
  });
}

export async function removeStoreItemTx({ itemId, warehouse, qty, reason, isMaterial, itemName }) {
  return runTransaction(db, async (tx) => {
    if (isMaterial) {
      const ref = doc(db, "materials", itemId);
      const snap = await tx.get(ref);
      const current = snap.data().stock || 0;
      if (current < qty) throw new Error(`Not enough ${itemName} in stock to remove.`);
      tx.update(ref, { stock: current - qty });
    } else {
      const ref = doc(db, "products", itemId);
      const snap = await tx.get(ref);
      const stockByBranch = snap.data().stockByBranch || {};
      const current = stockByBranch[warehouse] || 0;
      if (current < qty) throw new Error(`Not enough ${itemName} in stock at that warehouse to remove.`);
      tx.update(ref, { stockByBranch: { ...stockByBranch, [warehouse]: current - qty } });
    }
    tx.set(doc(collection(db, "movements")), {
      type: "OUT", itemType: isMaterial ? "Material" : "Product", item: itemId, qty,
      warehouse, date: todayStr(), ref: reason || "Manual Out",
    });
  });
}
