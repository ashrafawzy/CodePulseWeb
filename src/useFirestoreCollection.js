// One hook, reused for every collection in the app. This is what makes
// migrating ~36 entities from in-memory arrays to Firestore tractable: each
// entity becomes one call to this hook instead of hand-written Firestore
// wiring per module. It seeds a collection from the app's existing seed
// data the first time it's ever empty (so the demo data you already had
// shows up in Firestore automatically), then keeps local state live via
// Firestore's realtime listener — meaning add/update/delete never need to
// manually touch local state; Firestore's own snapshot does it for you.
import { useEffect, useState, useCallback } from "react";
import {
  collection, doc, onSnapshot, addDoc, setDoc, updateDoc, deleteDoc, getDocs, writeBatch,
} from "firebase/firestore";
import { db } from "./firebase";

export function useFirestoreCollection(name, seedData) {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  useEffect(() => {
    let unsub = () => {};
    (async () => {
      try {
        const colRef = collection(db, name);
        // One-time seed: only runs if the collection is genuinely empty,
        // so this never overwrites real data on later loads.
        const snap = await getDocs(colRef);
        if (snap.empty && seedData && seedData.length) {
          const batch = writeBatch(db);
          seedData.forEach((item) => {
            const { id, ...rest } = item;
            batch.set(doc(db, name, id), rest);
          });
          await batch.commit();
        }
        unsub = onSnapshot(
          colRef,
          (qs) => {
            setData(qs.docs.map((d) => ({ id: d.id, ...d.data() })));
            setLoading(false);
          },
          (err) => { setError(err); setLoading(false); }
        );
      } catch (err) {
        setError(err);
        setLoading(false);
      }
    })();
    return () => unsub();
  }, [name]);

  const add = useCallback(async (item, explicitId) => {
    if (explicitId) {
      await setDoc(doc(db, name, explicitId), item);
      return explicitId;
    }
    const ref = await addDoc(collection(db, name), item);
    return ref.id;
  }, [name]);

  const update = useCallback(async (id, patch) => {
    await updateDoc(doc(db, name, id), patch);
  }, [name]);

  const remove = useCallback(async (id) => {
    await deleteDoc(doc(db, name, id));
  }, [name]);

  return { data, loading, error, add, update, remove };
}
