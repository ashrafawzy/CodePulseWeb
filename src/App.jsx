import React, { useState, useMemo, useContext, createContext, useRef } from "react";
import * as XLSX from "xlsx";
import { useFirestoreCollection } from "./useFirestoreCollection";
import {
  uid, todayStr, receivePurchaseOrder, completeProductionBatch, transferProductStock,
  deliverSalesOrder, restockReturn, convertPurchaseRequestToOrder, convertQuotationToOrder,
  addStoreItemTx, removeStoreItemTx,
} from "./firestoreActions";

/* ============================================================
   CODEPULSE — Enterprise Operations Console
   Full module tree: Dashboard, CRM, Sales, Purchasing, Inventory,
   Accounting, HR, Expenses, Fixed Assets, Manufacturing, Projects,
   Service, Documents, Reports, Approvals, Users & Permissions, Settings.

   Data layer: Firebase Firestore (see useFirestoreCollection.js and
   firestoreActions.js) — every collection is real-time and shared across
   whoever has this deployed, not per-browser-session state anymore.
   ============================================================ */

/* ---------------- Themes ----------------
   A small set of design tokens exposed as CSS custom properties on the root
   element. Shared components (Card, Table, Button, SectionHeader, AddForm,
   Gauge) reference these via var(--token) instead of hardcoded hex, so
   switching themes re-colors the entire app without touching every page.
   The sidebar, top-bar brand mark, and Login screen intentionally stay
   fixed/branded in both themes — a common pattern (VS Code, Slack, etc.)
   that keeps the product identity consistent while the workspace adapts. */
const THEME_TOKENS = {
  light: {
    "--bg-page": "#F3EFE6",
    "--surface": "#FFFFFF",
    "--surface-alt": "#FBFAF6",
    "--border": "#E4E0D4",
    "--border-strong": "#C7C2B2",
    "--text": "#23271F",
    "--text-secondary": "#5C6B66",
    "--text-label": "#8A8578",
    "--heading": "#1B2421",
    "--zebra-border": "#EFEBE0",
  },
  dark: {
    "--bg-page": "#14170F",
    "--surface": "#1F241A",
    "--surface-alt": "#191D15",
    "--border": "#333B2A",
    "--border-strong": "#454E3A",
    "--text": "#E9E6DA",
    "--text-secondary": "#AEB39F",
    "--text-label": "#8B9080",
    "--heading": "#F3EFE6",
    "--zebra-border": "#2A3023",
  },
};

/* ---------------- Currencies ---------------- */
const seedCurrencies = [
  { code: "USD", rate: 1 },
  { code: "EUR", rate: 0.92 },
  { code: "GBP", rate: 0.79 },
  { code: "AED", rate: 3.67 },
  { code: "INR", rate: 83.1 },
  { code: "JPY", rate: 149 },
  { code: "EGP", rate: 49.5 }, // approximate — EGP has moved a lot; update the rate in Settings → Currencies
];
function formatMoney(amountUSD, code, rates) {
  const rate = rates?.[code] ?? 1;
  const val = amountUSD * rate;
  return val.toLocaleString("en-US", { style: "currency", currency: code, maximumFractionDigits: code === "JPY" ? 0 : 2 });
}

/* ---------------- Translations ---------------- */
const LANGUAGES = [
  { code: "en", label: "English" },
  { code: "es", label: "Español" },
  { code: "ar", label: "العربية" },
];
const I18N = {
  es: {
  "Dashboard": "Panel Principal",
  "CRM": "CRM",
  "Leads": "Prospectos",
  "Customers": "Clientes",
  "Opportunities": "Oportunidades",
  "Sales": "Ventas",
  "Quotations": "Cotizaciones",
  "Sales Orders": "Pedidos de Venta",
  "Sales Order": "Pedido de Venta",
  "Delivery": "Entrega",
  "Invoices": "Facturas",
  "Invoice": "Factura",
  "Returns": "Devoluciones",
  "Return": "Devolución",
  "Purchasing": "Compras",
  "Suppliers": "Proveedores",
  "Purchase Requests": "Solicitudes de Compra",
  "Purchase Orders": "Órdenes de Compra",
  "Receipts": "Recibos",
  "Receipt": "Recibo",
  "Inventory": "Inventario",
  "Items": "Artículos",
  "Item": "Artículo",
  "Warehouses": "Almacenes",
  "Warehouse": "Almacén",
  "Stock In": "Entrada de Stock",
  "Stock Out": "Salida de Stock",
  "Transfers": "Transferencias",
  "Transfer": "Transferencia",
  "Stock Count": "Conteo de Stock",
  "Store Management": "Gestión de Tiendas",
  "Add Items": "Agregar Artículos",
  "Out Items": "Retirar Artículos",
  "Item Limits": "Límites de Artículos",
  "Accounting": "Contabilidad",
  "Chart of Accounts": "Plan de Cuentas",
  "Journal Entries": "Asientos Contables",
  "Cash": "Caja",
  "Banks": "Bancos",
  "Payments": "Pagos",
  "Financial Reports": "Informes Financieros",
  "HR": "Recursos Humanos",
  "Employees": "Empleados",
  "Employee": "Empleado",
  "Attendance": "Asistencia",
  "Leaves": "Permisos",
  "Payroll": "Nómina",
  "Recruitment": "Reclutamiento",
  "Expenses": "Gastos",
  "Fixed Assets": "Activos Fijos",
  "Manufacturing": "Manufactura",
  "Production Orders": "Órdenes de Producción",
  "Work Centers": "Centros de Trabajo",
  "Quality Control": "Control de Calidad",
  "Maintenance": "Mantenimiento",
  "Work Center": "Centro de Trabajo",
  "Capacity / Day": "Capacidad / Día",
  "Devices": "Dispositivos",
  "Idle": "Inactivo",
  "+ Add Work Center": "+ Agregar Centro de Trabajo",
  "Inspections Logged": "Inspecciones Registradas",
  "Pass Rate": "Tasa de Aprobación",
  "Pending Review": "Pendiente de Revisión",
  "Inspector": "Inspector",
  "Result": "Resultado",
  "Notes": "Notas",
  "Pass": "Aprobado",
  "Fail": "Rechazado",
  "+ Log Inspection": "+ Registrar Inspección",
  "Equipment": "Equipo",
  "Last Service": "Último Servicio",
  "Next Due": "Próximo Vencimiento",
  "Technician": "Técnico",
  "Equipment Tracked": "Equipos Monitoreados",
  "Due Soon": "Próximo a Vencer",
  "Overdue": "Vencido",
  "Up to Date": "Al Día",
  "Needs attention": "Requiere Atención",
  "All clear": "Todo en Orden",
  "+ Add Equipment": "+ Agregar Equipo",
  "Log Service": "Registrar Servicio",
  "Projects": "Proyectos",
  "Project": "Proyecto",
  "Service": "Servicio",
  "Documents": "Documentos",
  "Document": "Documento",
  "Reports": "Informes",
  "Approvals": "Aprobaciones",
  "Users & Permissions": "Usuarios y Permisos",
  "Settings": "Configuración",
  "General": "General",
  "Currencies": "Monedas",
  "Enterprise snapshot across CRM, sales, production, and finance.": "Panorama general de CRM, ventas, producción y finanzas.",
  "Prospective hotel accounts not yet converted to customers.": "Cuentas de hoteles potenciales aún no convertidas en clientes.",
  "Hotel accounts with active or historical orders.": "Cuentas de hoteles con pedidos activos o históricos.",
  "Deals in progress with existing or prospective hotel accounts.": "Negocios en curso con cuentas de hoteles existentes o potenciales.",
  "Price quotes sent to hotels before a firm order is placed.": "Cotizaciones enviadas a los hoteles antes de confirmar un pedido.",
  "Confirmed orders from hotel clients, from intake through payment.": "Pedidos confirmados de clientes hoteleros, desde la entrada hasta el pago.",
  "Ship confirmed orders from their servicing branch.": "Envíe los pedidos confirmados desde su sucursal de servicio.",
  "Billing status for every delivered hotel order.": "Estado de facturación de cada pedido entregado a los hoteles.",
  "Return requests against delivered orders.": "Solicitudes de devolución de pedidos entregados.",
  "Vendors providing raw materials to the factory.": "Proveedores de materias primas para la fábrica.",
  "Internal requests for raw materials, approved before a PO is raised.": "Solicitudes internas de materias primas, aprobadas antes de emitir una orden de compra.",
  "Confirmed orders sent to suppliers for raw materials.": "Órdenes confirmadas enviadas a los proveedores para materias primas.",
  "Goods receipt notes for materials received into the factory.": "Notas de recepción de materiales ingresados a la fábrica.",
  "Supplier invoices generated when goods are received.": "Facturas de proveedores generadas al recibir la mercancía.",
  "Unified catalog of raw materials and finished goods.": "Catálogo unificado de materias primas y productos terminados.",
  "Factory and store locations holding inventory.": "Ubicaciones de fábrica y tiendas que contienen inventario.",
  "Move finished goods between warehouses.": "Mover productos terminados entre almacenes.",
  "Physical inventory counts reconciled against system quantities.": "Conteos físicos de inventario conciliados con las cantidades del sistema.",
  "Bring stock into a warehouse directly — new intake, found stock, or corrections.": "Ingrese existencias directamente a un almacén: nueva recepción, stock encontrado o correcciones.",
  "Remove stock from a warehouse for reasons outside a normal sale — damage, loss, samples, internal use.": "Retire existencias de un almacén por razones ajenas a una venta normal: daño, pérdida, muestras, uso interno.",
  "Set a minimum and maximum stock level per item, per warehouse.": "Establezca un nivel mínimo y máximo de existencias por artículo y por almacén.",
  "The ledger accounts used across journal entries and reports.": "Las cuentas contables utilizadas en los asientos e informes.",
  "Manual double-entry postings.": "Asientos manuales de partida doble.",
  "Simplified cash ledger — opening balance $15,000.": "Libro de caja simplificado: saldo inicial de $15,000.",
  "Company bank accounts.": "Cuentas bancarias de la empresa.",
  "Customer payments received.": "Pagos recibidos de clientes.",
  "Payments made to suppliers and reimbursed expenses.": "Pagos realizados a proveedores y gastos reembolsados.",
  "Simplified profit & loss summary — not a full accrual statement.": "Resumen simplificado de pérdidas y ganancias, no un estado contable completo.",
  "Factory and store staff.": "Personal de fábrica y tiendas.",
  "Daily attendance log.": "Registro diario de asistencia.",
  "Employee leave requests.": "Solicitudes de permiso de empleados.",
  "Monthly payroll run.": "Nómina mensual.",
  "Open positions and candidate pipeline.": "Vacantes abiertas y proceso de candidatos.",
  "Employee expense claims.": "Reclamos de gastos de empleados.",
  "Asset register with straight-line depreciation.": "Registro de activos con depreciación en línea recta.",
  "Plan batches at the factory, track material consumption, and complete runs.": "Planifique lotes en la fábrica, controle el consumo de materiales y complete las corridas.",
  "Internal and client-facing initiatives.": "Iniciativas internas y de cara al cliente.",
  "Post-sale support tickets from hotel clients.": "Tickets de soporte posventa de clientes hoteleros.",
  "Contracts, reports, and compliance records.": "Contratos, informes y registros de cumplimiento.",
  "Cross-module summary. Deeper report builders can be added as this system grows.": "Resumen entre módulos. Se pueden añadir generadores de informes más avanzados a medida que crezca el sistema.",
  "Everything across the system waiting on a decision.": "Todo lo que en el sistema espera una decisión.",
  "Accounts and the modules each role can access.": "Cuentas y los módulos a los que puede acceder cada rol.",
  "Company-wide configuration (demo only — not persisted).": "Configuración de toda la empresa (solo demostración, no se guarda).",
  "Exchange rates used everywhere money is shown or entered across the system. Rates are against USD.": "Tipos de cambio utilizados en todo lugar donde se muestra o ingresa dinero en el sistema. Las tasas son respecto al USD.",
  "ID": "ID",
  "Date": "Fecha",
  "Status": "Estado",
  "Amount": "Monto",
  "Total": "Total",
  "Qty": "Cant.",
  "Name": "Nombre",
  "Email": "Correo",
  "Type": "Tipo",
  "Category": "Categoría",
  "Contact": "Contacto",
  "Location": "Ubicación",
  "Branch": "Sucursal",
  "Department": "Departamento",
  "Role": "Rol",
  "Priority": "Prioridad",
  "Reason": "Motivo",
  "Client": "Cliente",
  "Customer": "Cliente",
  "Supplier": "Proveedor",
  "Material": "Material",
  "Product": "Producto",
  "Reference": "Referencia",
  "Description": "Descripción",
  "Memo": "Memo",
  "Code": "Código",
  "Account": "Cuenta",
  "Account Name": "Nombre de la Cuenta",
  "Balance": "Saldo",
  "Rate vs USD": "Tasa frente al USD",
  "Currently Displaying": "Mostrando Actualmente",
  "Min": "Mín.",
  "Max": "Máx.",
  "Stock": "Existencias",
  "Current Stock": "Existencia Actual",
  "Reorder At": "Reordenar en",
  "Unit Value": "Valor Unitario",
  "Cost": "Costo",
  "Deal": "Negocio",
  "Stage": "Etapa",
  "Value": "Valor",
  "Close Date": "Fecha de Cierre",
  "Requested By": "Solicitado Por",
  "Assigned To": "Asignado A",
  "Issue": "Problema",
  "Manager": "Gerente",
  "Budget": "Presupuesto",
  "Progress": "Progreso",
  "Position": "Puesto",
  "Candidate": "Candidato",
  "Month": "Mes",
  "Base": "Base",
  "Deductions": "Deducciones",
  "Net Pay": "Pago Neto",
  "Hired": "Contratado",
  "Lead Time": "Tiempo de Entrega",
  "System Qty": "Cant. del Sistema",
  "Counted Qty": "Cant. Contada",
  "Variance": "Variación",
  "Number": "Número",
  "Bank Account": "Cuenta Bancaria",
  "Payee": "Beneficiario",
  "Ref": "Ref.",
  "Uploaded By": "Subido Por",
  "Useful Life": "Vida Útil",
  "Book Value": "Valor en Libros",
  "Purchase Date": "Fecha de Compra",
  "PR": "SC",
  "PO": "OC",
  "Batch": "Lote",
  "Materials Needed": "Materiales Necesarios",
  "Servicing Branch": "Sucursal de Servicio",
  "Lifetime Orders": "Pedidos Totales",
  "Lifetime Value": "Valor Total",
  "From": "Desde",
  "To": "Hasta",
  "Group": "Grupo",
  "Source": "Origen",
  "Owner": "Responsable",
  "Module": "Módulo",
  "Entry": "Asiento",
  "Debit": "Débito",
  "Credit": "Crédito",
  "Debit Account": "Cuenta Deudora",
  "Credit Account": "Cuenta Acreedora",
  "Inflow": "Entrada",
  "Outflow": "Salida",
  "Order": "Pedido",
  "Quote": "Cotización",
  "Ticket": "Ticket",
  "Run": "Ejecución",
  "Movement": "Movimiento",
  "Count": "Conteo",
  "Hotel": "Hotel",
  "Hotel Name": "Nombre del Hotel",
  "Requester": "Solicitante",
  "Detail": "Detalle",
  "HQ Factory": "Fábrica Central",
  "Denver Store": "Tienda Denver",
  "Tampa Store": "Tienda Tampa",
  "Username": "Usuario",
  "Code (e.g. CAD)": "Código (ej. CAD)",
  "Pending": "Pendiente",
  "Fulfilled": "Cumplido",
  "Delivered": "Entregado",
  "Invoiced": "Facturado",
  "Paid": "Pagado",
  "Unpaid": "Sin Pagar",
  "Planned": "Planificado",
  "Completed": "Completado",
  "Received": "Recibido",
  "Draft": "Borrador",
  "Sent": "Enviado",
  "Accepted": "Aceptado",
  "Rejected": "Rechazado",
  "Converted": "Convertido",
  "Approved": "Aprobado",
  "Requested": "Solicitado",
  "Restocked": "Reabastecido",
  "New": "Nuevo",
  "Contacted": "Contactado",
  "Qualified": "Calificado",
  "Lost": "Perdido",
  "Prospecting": "Prospección",
  "Proposal": "Propuesta",
  "Negotiation": "Negociación",
  "Won": "Ganado",
  "Open": "Abierto",
  "In Progress": "En Progreso",
  "Resolved": "Resuelto",
  "Active": "Activo",
  "Inactive": "Inactivo",
  "Taxes": "Impuestos",
  "Tax Name": "Nombre del Impuesto",
  "Rate %": "Tasa %",
  "Applies To": "Aplica A",
  "Sales Tax": "Impuesto sobre Ventas",
  "VAT": "IVA",
  "Duty": "Arancel",
  "Withholding": "Retención",
  "Activate": "Activar",
  "Deactivate": "Desactivar",
  "+ Add Tax Rate": "+ Agregar Tasa de Impuesto",
  "Tax Rates Configured": "Tasas de Impuesto Configuradas",
  "Active Rates": "Tasas Activas",
  "Average Rate": "Tasa Promedio",
  "Active Batches": "Lotes Activos",
  "Units In Production": "Unidades en Producción",
  "Completed Batches": "Lotes Completados",
  "Batches At Risk": "Lotes en Riesgo",
  "Applied": "Aplicado",
  "Interview": "Entrevista",
  "Offer": "Oferta",
  "Planning": "Planificación",
  "On Hold": "En Espera",
  "Present": "Presente",
  "Absent": "Ausente",
  "Late": "Tarde",
  "Raw Material": "Materia Prima",
  "Finished Good": "Producto Terminado",
  "Expense": "Gasto",
  "Purchase Request": "Solicitud de Compra",
  "Leave Request": "Solicitud de Permiso",
  "Expense Claim": "Reclamo de Gasto",
  "Below Min": "Bajo el Mínimo",
  "Within Range": "Dentro del Rango",
  "Above Max": "Sobre el Máximo",
  "Sick": "Enfermedad",
  "Vacation": "Vacaciones",
  "Personal": "Personal",
  "Damaged": "Dañado",
  "Sample": "Muestra",
  "Internal Use": "Uso Interno",
  "Other": "Otro",
  "Low": "Baja",
  "Medium": "Media",
  "High": "Alta",
  "Asset": "Activo",
  "Liability": "Pasivo",
  "Equity": "Patrimonio",
  "Revenue": "Ingreso",
  "January": "Enero",
  "April": "Abril",
  "July": "Julio",
  "October": "Octubre",
  "+ Add": "+ Agregar",
  "+ Add Currency": "+ Agregar Moneda",
  "+ Add Stock": "+ Agregar Existencias",
  "− Remove Stock": "− Retirar Existencias",
  "+ Create Quotation": "+ Crear Cotización",
  "+ Create PO": "+ Crear Orden de Compra",
  "+ New Project": "+ Nuevo Proyecto",
  "+ New Ticket": "+ Nuevo Ticket",
  "+ Record Count": "+ Registrar Conteo",
  "+ Request": "+ Solicitar",
  "+ Request Leave": "+ Solicitar Permiso",
  "+ Request Return": "+ Solicitar Devolución",
  "+ Submit Claim": "+ Enviar Reclamo",
  "+ Schedule Batch": "+ Programar Lote",
  "+ Upload Document": "+ Subir Documento",
  "Set Limit": "Establecer Límite",
  "Transfer Stock": "Transferir Existencias",
  "Update Rate": "Actualizar Tasa",
  "Save Settings": "Guardar Configuración",
  "Delete": "Eliminar",
  "Edit": "Editar",
  "Save Changes": "Guardar Cambios",
  "+ Add Supplier": "+ Agregar Proveedor",
  "+ Add User": "+ Agregar Usuario",
  "+ Add Bank Account": "+ Agregar Cuenta Bancaria",
  "+ Add Asset": "+ Agregar Activo",
  "+ Add Document": "+ Agregar Documento",
  "+ Record Receipt": "+ Registrar Recibo",
  "+ Record Payment": "+ Registrar Pago",
  "Approve": "Aprobar",
  "Reject": "Rechazar",
  "Receive": "Recibir",
  "Send": "Enviar",
  "Mark Accepted": "Marcar como Aceptado",
  "Convert to Order": "Convertir en Pedido",
  "Convert to PO": "Convertir en Orden de Compra",
  "Mark Delivered": "Marcar como Entregado",
  "Create Invoice": "Crear Factura",
  "Mark Paid": "Marcar como Pagado",
  "Restock": "Reabastecer",
  "Complete": "Completar",
  "Short stock": "Existencia Insuficiente",
  "Start Work": "Iniciar Trabajo",
  "Mark Resolved": "Marcar como Resuelto",
  "Enterprise Console": "Consola Empresarial",
  "Enterprise Console — Sign In": "Consola Empresarial — Iniciar Sesión",
  "Password (demo — any value works)": "Contraseña (demo — cualquier valor funciona)",
  "Sign In": "Iniciar Sesión",
  "Sign Out": "Cerrar Sesión",
  "Modules and actions adapt to the signed-in role — Viewer is read-only everywhere.": "Los módulos y las acciones se adaptan al rol conectado; el rol Visor es de solo lectura en todas partes.",
  "Line status:": "Estado de la línea:",
  "Running": "En Funcionamiento",
  "Currency": "Moneda",
  "Theme": "Tema",
  "Light": "Claro",
  "Dark": "Oscuro",
  "Language": "Idioma",
  "Admin": "Administrador",
  "Factory Manager": "Gerente de Fábrica",
  "Sales Manager": "Gerente de Ventas",
  "Accountant": "Contador",
  "Viewer": "Visor",
  },
  ar: {
  "Dashboard": "لوحة التحكم",
  "CRM": "إدارة علاقات العملاء",
  "Leads": "العملاء المحتملون",
  "Customers": "العملاء",
  "Opportunities": "الفرص",
  "Sales": "المبيعات",
  "Quotations": "عروض الأسعار",
  "Sales Orders": "أوامر البيع",
  "Sales Order": "أمر بيع",
  "Delivery": "التسليم",
  "Invoices": "الفواتير",
  "Invoice": "فاتورة",
  "Returns": "المرتجعات",
  "Return": "مرتجع",
  "Purchasing": "المشتريات",
  "Suppliers": "الموردون",
  "Purchase Requests": "طلبات الشراء",
  "Purchase Orders": "أوامر الشراء",
  "Receipts": "الإيصالات",
  "Receipt": "إيصال",
  "Inventory": "المخزون",
  "Items": "الأصناف",
  "Item": "صنف",
  "Warehouses": "المستودعات",
  "Warehouse": "المستودع",
  "Stock In": "إدخال المخزون",
  "Stock Out": "إخراج المخزون",
  "Transfers": "التحويلات",
  "Transfer": "تحويل",
  "Stock Count": "جرد المخزون",
  "Store Management": "إدارة المتاجر",
  "Add Items": "إضافة أصناف",
  "Out Items": "إخراج أصناف",
  "Item Limits": "حدود الأصناف",
  "Accounting": "المحاسبة",
  "Chart of Accounts": "دليل الحسابات",
  "Journal Entries": "قيود اليومية",
  "Cash": "النقدية",
  "Banks": "البنوك",
  "Payments": "المدفوعات",
  "Financial Reports": "التقارير المالية",
  "HR": "الموارد البشرية",
  "Employees": "الموظفون",
  "Employee": "موظف",
  "Attendance": "الحضور",
  "Leaves": "الإجازات",
  "Payroll": "كشف الرواتب",
  "Recruitment": "التوظيف",
  "Expenses": "المصروفات",
  "Fixed Assets": "الأصول الثابتة",
  "Manufacturing": "التصنيع",
  "Production Orders": "أوامر الإنتاج",
  "Work Centers": "مراكز العمل",
  "Quality Control": "مراقبة الجودة",
  "Maintenance": "الصيانة",
  "Work Center": "مركز العمل",
  "Capacity / Day": "الطاقة الإنتاجية / يوم",
  "Devices": "الأجهزة",
  "Idle": "خامل",
  "+ Add Work Center": "+ إضافة مركز عمل",
  "Inspections Logged": "الفحوصات المسجَّلة",
  "Pass Rate": "نسبة النجاح",
  "Pending Review": "قيد المراجعة",
  "Inspector": "المفتش",
  "Result": "النتيجة",
  "Notes": "ملاحظات",
  "Pass": "ناجح",
  "Fail": "راسب",
  "+ Log Inspection": "+ تسجيل فحص",
  "Equipment": "المعدات",
  "Last Service": "آخر صيانة",
  "Next Due": "الموعد القادم",
  "Technician": "الفني",
  "Equipment Tracked": "المعدات المتابَعة",
  "Due Soon": "يستحق قريبًا",
  "Overdue": "متأخر",
  "Up to Date": "محدَّث",
  "Needs attention": "يحتاج انتباه",
  "All clear": "كل شيء على ما يرام",
  "+ Add Equipment": "+ إضافة معدة",
  "Log Service": "تسجيل الصيانة",
  "Projects": "المشاريع",
  "Project": "مشروع",
  "Service": "الخدمة",
  "Documents": "المستندات",
  "Document": "مستند",
  "Reports": "التقارير",
  "Approvals": "الموافقات",
  "Users & Permissions": "المستخدمون والصلاحيات",
  "Settings": "الإعدادات",
  "General": "عام",
  "Currencies": "العملات",
  "Enterprise snapshot across CRM, sales, production, and finance.": "لمحة شاملة عن إدارة العملاء والمبيعات والإنتاج والتمويل.",
  "Prospective hotel accounts not yet converted to customers.": "حسابات فنادق محتملة لم يتم تحويلها بعد إلى عملاء.",
  "Hotel accounts with active or historical orders.": "حسابات فنادق لديها طلبات حالية أو سابقة.",
  "Deals in progress with existing or prospective hotel accounts.": "صفقات جارية مع حسابات فنادق حالية أو محتملة.",
  "Price quotes sent to hotels before a firm order is placed.": "عروض أسعار مرسلة إلى الفنادق قبل تأكيد الطلب.",
  "Confirmed orders from hotel clients, from intake through payment.": "طلبات مؤكدة من عملاء الفنادق، من الاستلام حتى الدفع.",
  "Ship confirmed orders from their servicing branch.": "شحن الطلبات المؤكدة من الفرع الذي يخدمها.",
  "Billing status for every delivered hotel order.": "حالة الفوترة لكل طلب فندقي تم تسليمه.",
  "Return requests against delivered orders.": "طلبات إرجاع للطلبات المُسلَّمة.",
  "Vendors providing raw materials to the factory.": "الموردون الذين يوفرون المواد الخام للمصنع.",
  "Internal requests for raw materials, approved before a PO is raised.": "طلبات داخلية للمواد الخام، تُعتمد قبل إصدار أمر الشراء.",
  "Confirmed orders sent to suppliers for raw materials.": "أوامر مؤكدة تُرسل إلى الموردين للحصول على المواد الخام.",
  "Goods receipt notes for materials received into the factory.": "إشعارات استلام البضائع للمواد الواردة إلى المصنع.",
  "Supplier invoices generated when goods are received.": "فواتير الموردين التي تُنشأ عند استلام البضائع.",
  "Unified catalog of raw materials and finished goods.": "كتالوج موحّد للمواد الخام والمنتجات النهائية.",
  "Factory and store locations holding inventory.": "مواقع المصنع والمتاجر التي تحتفظ بالمخزون.",
  "Move finished goods between warehouses.": "نقل المنتجات النهائية بين المستودعات.",
  "Physical inventory counts reconciled against system quantities.": "جرد مادي للمخزون تتم مطابقته مع كميات النظام.",
  "Bring stock into a warehouse directly — new intake, found stock, or corrections.": "إدخال المخزون إلى المستودع مباشرة — استلام جديد، مخزون تم العثور عليه، أو تصحيحات.",
  "Remove stock from a warehouse for reasons outside a normal sale — damage, loss, samples, internal use.": "إخراج المخزون من المستودع لأسباب خارج نطاق البيع العادي — تلف، فقدان، عينات، استخدام داخلي.",
  "Set a minimum and maximum stock level per item, per warehouse.": "تحديد حد أدنى وأقصى للمخزون لكل صنف ولكل مستودع.",
  "The ledger accounts used across journal entries and reports.": "الحسابات الدفترية المستخدمة في القيود والتقارير.",
  "Manual double-entry postings.": "قيود يدوية بنظام القيد المزدوج.",
  "Simplified cash ledger — opening balance $15,000.": "دفتر نقدية مبسّط — الرصيد الافتتاحي 15,000 دولار.",
  "Company bank accounts.": "الحسابات المصرفية للشركة.",
  "Customer payments received.": "المدفوعات المستلمة من العملاء.",
  "Payments made to suppliers and reimbursed expenses.": "المدفوعات إلى الموردين والمصروفات المستردة.",
  "Simplified profit & loss summary — not a full accrual statement.": "ملخص مبسّط للأرباح والخسائر — وليس بيانًا محاسبيًا كاملاً.",
  "Factory and store staff.": "موظفو المصنع والمتاجر.",
  "Daily attendance log.": "سجل الحضور اليومي.",
  "Employee leave requests.": "طلبات إجازة الموظفين.",
  "Monthly payroll run.": "تشغيل كشف الرواتب الشهري.",
  "Open positions and candidate pipeline.": "الوظائف الشاغرة ومسار المرشحين.",
  "Employee expense claims.": "مطالبات مصروفات الموظفين.",
  "Asset register with straight-line depreciation.": "سجل الأصول مع الإهلاك بطريقة القسط الثابت.",
  "Plan batches at the factory, track material consumption, and complete runs.": "خطط للدفعات في المصنع، وتتبّع استهلاك المواد، وأكمل عمليات التشغيل.",
  "Internal and client-facing initiatives.": "مبادرات داخلية وموجهة للعملاء.",
  "Post-sale support tickets from hotel clients.": "تذاكر الدعم بعد البيع من عملاء الفنادق.",
  "Contracts, reports, and compliance records.": "العقود والتقارير وسجلات الامتثال.",
  "Cross-module summary. Deeper report builders can be added as this system grows.": "ملخص شامل بين الوحدات. يمكن إضافة أدوات تقارير أكثر تفصيلاً مع نمو النظام.",
  "Everything across the system waiting on a decision.": "كل ما ينتظر قرارًا في النظام.",
  "Accounts and the modules each role can access.": "الحسابات والوحدات التي يمكن لكل دور الوصول إليها.",
  "Company-wide configuration (demo only — not persisted).": "إعدادات على مستوى الشركة (تجريبي فقط — لا يتم حفظه).",
  "Exchange rates used everywhere money is shown or entered across the system. Rates are against USD.": "أسعار الصرف المستخدمة في كل مكان يُعرض أو يُدخل فيه المال في النظام. الأسعار مقابل الدولار الأمريكي.",
  "ID": "المعرف",
  "Date": "التاريخ",
  "Status": "الحالة",
  "Amount": "المبلغ",
  "Total": "الإجمالي",
  "Qty": "الكمية",
  "Name": "الاسم",
  "Email": "البريد الإلكتروني",
  "Type": "النوع",
  "Category": "الفئة",
  "Contact": "جهة الاتصال",
  "Location": "الموقع",
  "Branch": "الفرع",
  "Department": "القسم",
  "Role": "الدور",
  "Priority": "الأولوية",
  "Reason": "السبب",
  "Client": "العميل",
  "Customer": "العميل",
  "Supplier": "المورّد",
  "Material": "المادة",
  "Product": "المنتج",
  "Reference": "المرجع",
  "Description": "الوصف",
  "Memo": "ملاحظة",
  "Code": "الرمز",
  "Account": "الحساب",
  "Account Name": "اسم الحساب",
  "Balance": "الرصيد",
  "Rate vs USD": "السعر مقابل الدولار",
  "Currently Displaying": "المعروض حاليًا",
  "Min": "الحد الأدنى",
  "Max": "الحد الأقصى",
  "Stock": "المخزون",
  "Current Stock": "المخزون الحالي",
  "Reorder At": "إعادة الطلب عند",
  "Unit Value": "قيمة الوحدة",
  "Cost": "التكلفة",
  "Deal": "الصفقة",
  "Stage": "المرحلة",
  "Value": "القيمة",
  "Close Date": "تاريخ الإغلاق",
  "Requested By": "طلب بواسطة",
  "Assigned To": "مُسند إلى",
  "Issue": "المشكلة",
  "Manager": "المدير",
  "Budget": "الميزانية",
  "Progress": "التقدم",
  "Position": "المنصب",
  "Candidate": "المرشح",
  "Month": "الشهر",
  "Base": "الأساسي",
  "Deductions": "الخصومات",
  "Net Pay": "صافي الراتب",
  "Hired": "تم التوظيف",
  "Lead Time": "مدة التوريد",
  "System Qty": "كمية النظام",
  "Counted Qty": "الكمية المعدودة",
  "Variance": "الفرق",
  "Number": "الرقم",
  "Bank Account": "الحساب المصرفي",
  "Payee": "المستفيد",
  "Ref": "المرجع",
  "Uploaded By": "تم الرفع بواسطة",
  "Useful Life": "العمر الإنتاجي",
  "Book Value": "القيمة الدفترية",
  "Purchase Date": "تاريخ الشراء",
  "PR": "طلب شراء",
  "PO": "أمر شراء",
  "Batch": "الدفعة",
  "Materials Needed": "المواد المطلوبة",
  "Servicing Branch": "فرع الخدمة",
  "Lifetime Orders": "إجمالي الطلبات",
  "Lifetime Value": "القيمة الإجمالية",
  "From": "من",
  "To": "إلى",
  "Group": "المجموعة",
  "Source": "المصدر",
  "Owner": "المسؤول",
  "Module": "الوحدة",
  "Entry": "القيد",
  "Debit": "مدين",
  "Credit": "دائن",
  "Debit Account": "الحساب المدين",
  "Credit Account": "الحساب الدائن",
  "Inflow": "الوارد",
  "Outflow": "الصادر",
  "Order": "الطلب",
  "Quote": "عرض السعر",
  "Ticket": "التذكرة",
  "Run": "التشغيل",
  "Movement": "الحركة",
  "Count": "الجرد",
  "Hotel": "الفندق",
  "Hotel Name": "اسم الفندق",
  "Requester": "مقدّم الطلب",
  "Detail": "التفاصيل",
  "HQ Factory": "المصنع الرئيسي",
  "Denver Store": "متجر دنفر",
  "Tampa Store": "متجر تامبا",
  "Username": "اسم المستخدم",
  "Code (e.g. CAD)": "الرمز (مثال CAD)",
  "Pending": "قيد الانتظار",
  "Fulfilled": "منفَّذ",
  "Delivered": "تم التسليم",
  "Invoiced": "تمت الفوترة",
  "Paid": "مدفوع",
  "Unpaid": "غير مدفوع",
  "Planned": "مخطَّط",
  "Completed": "مكتمل",
  "Received": "تم الاستلام",
  "Draft": "مسودة",
  "Sent": "تم الإرسال",
  "Accepted": "مقبول",
  "Rejected": "مرفوض",
  "Converted": "تم التحويل",
  "Approved": "معتمد",
  "Requested": "تم الطلب",
  "Restocked": "تمت إعادة التخزين",
  "New": "جديد",
  "Contacted": "تم الاتصال",
  "Qualified": "مؤهَّل",
  "Lost": "خسارة",
  "Prospecting": "استكشاف",
  "Proposal": "عرض",
  "Negotiation": "تفاوض",
  "Won": "فوز",
  "Open": "مفتوح",
  "In Progress": "قيد التنفيذ",
  "Resolved": "تم الحل",
  "Active": "نشط",
  "Inactive": "غير نشط",
  "Taxes": "الضرائب",
  "Tax Name": "اسم الضريبة",
  "Rate %": "النسبة %",
  "Applies To": "تُطبَّق على",
  "Sales Tax": "ضريبة المبيعات",
  "VAT": "ضريبة القيمة المضافة",
  "Duty": "رسوم جمركية",
  "Withholding": "ضريبة الاستقطاع",
  "Activate": "تفعيل",
  "Deactivate": "إلغاء التفعيل",
  "+ Add Tax Rate": "+ إضافة نسبة ضريبة",
  "Tax Rates Configured": "نسب الضرائب المُعدَّة",
  "Active Rates": "النسب النشطة",
  "Average Rate": "متوسط النسبة",
  "Active Batches": "الدفعات النشطة",
  "Units In Production": "الوحدات قيد الإنتاج",
  "Completed Batches": "الدفعات المكتملة",
  "Batches At Risk": "الدفعات المعرَّضة للخطر",
  "Applied": "تم التقديم",
  "Interview": "مقابلة",
  "Offer": "عرض عمل",
  "Planning": "التخطيط",
  "On Hold": "معلَّق",
  "Present": "حاضر",
  "Absent": "غائب",
  "Late": "متأخر",
  "Raw Material": "مادة خام",
  "Finished Good": "منتج نهائي",
  "Expense": "مصروف",
  "Purchase Request": "طلب شراء",
  "Leave Request": "طلب إجازة",
  "Expense Claim": "مطالبة مصروف",
  "Below Min": "أقل من الحد الأدنى",
  "Within Range": "ضمن النطاق",
  "Above Max": "أعلى من الحد الأقصى",
  "Sick": "مرضية",
  "Vacation": "إجازة سنوية",
  "Personal": "شخصية",
  "Damaged": "تالف",
  "Sample": "عينة",
  "Internal Use": "استخدام داخلي",
  "Other": "أخرى",
  "Low": "منخفضة",
  "Medium": "متوسطة",
  "High": "عالية",
  "Asset": "أصل",
  "Liability": "خصم",
  "Equity": "حقوق الملكية",
  "Revenue": "إيراد",
  "January": "يناير",
  "April": "أبريل",
  "July": "يوليو",
  "October": "أكتوبر",
  "+ Add": "+ إضافة",
  "+ Add Currency": "+ إضافة عملة",
  "+ Add Stock": "+ إضافة مخزون",
  "− Remove Stock": "− إخراج مخزون",
  "+ Create Quotation": "+ إنشاء عرض سعر",
  "+ Create PO": "+ إنشاء أمر شراء",
  "+ New Project": "+ مشروع جديد",
  "+ New Ticket": "+ تذكرة جديدة",
  "+ Record Count": "+ تسجيل جرد",
  "+ Request": "+ طلب",
  "+ Request Leave": "+ طلب إجازة",
  "+ Request Return": "+ طلب إرجاع",
  "+ Submit Claim": "+ إرسال مطالبة",
  "+ Schedule Batch": "+ جدولة دفعة",
  "+ Upload Document": "+ رفع مستند",
  "Set Limit": "تعيين الحد",
  "Transfer Stock": "تحويل المخزون",
  "Update Rate": "تحديث السعر",
  "Save Settings": "حفظ الإعدادات",
  "Delete": "حذف",
  "Edit": "تعديل",
  "Save Changes": "حفظ التغييرات",
  "+ Add Supplier": "+ إضافة مورد",
  "+ Add User": "+ إضافة مستخدم",
  "+ Add Bank Account": "+ إضافة حساب بنكي",
  "+ Add Asset": "+ إضافة أصل",
  "+ Add Document": "+ إضافة مستند",
  "+ Record Receipt": "+ تسجيل إيصال",
  "+ Record Payment": "+ تسجيل دفعة",
  "Approve": "اعتماد",
  "Reject": "رفض",
  "Receive": "استلام",
  "Send": "إرسال",
  "Mark Accepted": "تمييز كمقبول",
  "Convert to Order": "تحويل إلى طلب",
  "Convert to PO": "تحويل إلى أمر شراء",
  "Mark Delivered": "تمييز كمُسلَّم",
  "Create Invoice": "إنشاء فاتورة",
  "Mark Paid": "تمييز كمدفوع",
  "Restock": "إعادة التخزين",
  "Complete": "إكمال",
  "Short stock": "المخزون غير كافٍ",
  "Start Work": "بدء العمل",
  "Mark Resolved": "تمييز كمحلول",
  "Enterprise Console": "لوحة التحكم المؤسسية",
  "Enterprise Console — Sign In": "لوحة التحكم المؤسسية — تسجيل الدخول",
  "Password (demo — any value works)": "كلمة المرور (تجريبي — أي قيمة تعمل)",
  "Sign In": "تسجيل الدخول",
  "Sign Out": "تسجيل الخروج",
  "Modules and actions adapt to the signed-in role — Viewer is read-only everywhere.": "تتكيّف الوحدات والإجراءات مع الدور المسجَّل — دور \"مشاهد\" للقراءة فقط في كل مكان.",
  "Line status:": "حالة الخط:",
  "Running": "قيد التشغيل",
  "Currency": "العملة",
  "Theme": "المظهر",
  "Light": "فاتح",
  "Dark": "داكن",
  "Language": "اللغة",
  "Admin": "مسؤول",
  "Factory Manager": "مدير المصنع",
  "Sales Manager": "مدير المبيعات",
  "Accountant": "محاسب",
  "Viewer": "مشاهد",
  },
};

const LOGO_DATA_URI = "data:image/png;base64,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Fm8EQVBVVbV9+/ann36aklJP0L5AYrcNGiuCGrTrOsOGDSO91O6L9Vat9dixY+fMmYOq4APcvZVAsPf+GuirK2l25MiRnHPSUxUrjLVKE+L60vUQQl555ZU333yzogkKOO3r0D00QfH46uqauro6rXWvbESMYh3HSZJk7ty51dXVxWJhgAegH45Vll8iQjCtdaaqaqBJAR6YxdXW1u7atWvx4r8iJBUEwbtlce9t0EJw7FmvHzIkk8kkcdxb9BkgBIkmRxxxxMknz+7Feu/7+ER//+vDsBACAzDYg+g4jpCix3ujzDRkpC9vkErUE088sXLlq0jERTSphzE0alEyxqpra0os9d6A2Cilgpe2SjqVmjNnLkrydPmU/X98d8egPwxmzTnHApAx2rYtz/XMoffg49FNKWG8LDjWe/6oy4gj7bru+g3r77//fmOMZdm+X0yn08YAOVgq+96mWRIGKPpTp0ypNAX21qc2SksugqB48uxZR0ycEEchp4QC1YmiQDjt5wiEvdcXKTNM/v7XQE8Ky6I/wDnVWqVSLmI175U8sC4vAlontsXjKKyrqa3KZJAl2Dt5CKWGaC5FGEeUs0TrPz3wwOYtW1PpbBQraTlRrLgQBN6P+igltm2XkKnevLl4f01VVfb000+3bStJEsaQZAxGkwFedftwBNU9iVuASM4IGMZEbW0No71W8KeEGm08L6W0YozV1dU99NBDixc/kU5ncUgupbQ0RYr0VMEfHbNtlYQdSO81CCHyg1Shs846s66uVggRJwrdt4HDz+wZXO/28BjniUpc1x4zZgySaXuHEwHGduxcLu8X/YaGYc888+y9996T68xxzrqz97prLowx23F63UOXk26IonDUqBFnnnlWZy4nOC8LtZhByxnAjh2CMKypqZ0wYaLWqsfSvQdGRKg+PmLEyM2bN9955x1bNm+tqanF1oSK2RNiei7WiAGW49h9UQxDATXsbLvwwoUNDUMp26f5dNB0BmQgXurAHTlyxPjx45HB0itRB2MsDMKqqqqmpqaf/OQna9e+7Tgumng306DuhPAl/cxevimsVGfCHqEgCEaNGn3BBRfk8zns7EWbHrSeAbuEEA0NDUOGDMGBGO9zxkilm6kqW1Us+nfeecfixU8QoIdEs+6OQZc+peu6OFGh97ymIdQIibAJkijIGWecPnz4cGVixqjR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";

/* ---------------- Branches / Warehouses ---------------- */
const BRANCHES = [
  { id: "BR-HQ", name: "Chicago Factory & Warehouse", type: "Factory", location: "Chicago, IL" },
  { id: "BR-DEN", name: "Denver Distribution Store", type: "Store", location: "Denver, CO" },
  { id: "BR-TPA", name: "Tampa Distribution Store", type: "Store", location: "Tampa, FL" },
];

/* ---------------- Users & Roles ---------------- */
const seedUsers = [
  { id: "u1", name: "Sam Rivera", role: "Admin", username: "admin" },
  { id: "u2", name: "Priya Nair", role: "Factory Manager", username: "factory" },
  { id: "u3", name: "Jordan Blake", role: "Sales Manager", username: "sales" },
  { id: "u4", name: "Casey Wu", role: "Accountant", username: "finance" },
  { id: "u5", name: "Morgan Lee", role: "Viewer", username: "viewer" },
];
const ROLES_LIST = ["Admin", "Factory Manager", "Sales Manager", "Accountant", "Viewer"];

const MODULES = [
  { id: "dashboard", label: "Dashboard" },
  { id: "crm", label: "CRM", subs: [{ id: "leads", label: "Leads" }, { id: "customers", label: "Customers" }, { id: "opportunities", label: "Opportunities" }] },
  { id: "sales", label: "Sales", subs: [{ id: "quotations", label: "Quotations" }, { id: "salesorders", label: "Sales Orders" }, { id: "delivery", label: "Delivery" }, { id: "invoices", label: "Invoices" }, { id: "returns", label: "Returns" }] },
  { id: "purchasing", label: "Purchasing", subs: [{ id: "suppliers", label: "Suppliers" }, { id: "purchaserequests", label: "Purchase Requests" }, { id: "purchaseorders", label: "Purchase Orders" }, { id: "receipts", label: "Receipts" }, { id: "purchaseinvoices", label: "Invoices" }] },
  { id: "inventory", label: "Inventory", subs: [{ id: "items", label: "Items" }, { id: "warehouses", label: "Warehouses" }, { id: "stockin", label: "Stock In" }, { id: "stockout", label: "Stock Out" }, { id: "transfers", label: "Transfers" }, { id: "stockcount", label: "Stock Count" }] },
  { id: "storemgmt", label: "Store Management", subs: [{ id: "additems", label: "Add Items" }, { id: "outitems", label: "Out Items" }, { id: "sm_transfers", label: "Transfer" }, { id: "itemlimits", label: "Item Limits" }] },
  { id: "accounting", label: "Accounting", subs: [{ id: "coa", label: "Chart of Accounts" }, { id: "journal", label: "Journal Entries" }, { id: "cash", label: "Cash" }, { id: "banks", label: "Banks" }, { id: "arreceipts", label: "Receipts" }, { id: "appayments", label: "Payments" }, { id: "taxes", label: "Taxes" }, { id: "finreports", label: "Financial Reports" }] },
  { id: "hr", label: "HR", subs: [{ id: "employees", label: "Employees" }, { id: "attendance", label: "Attendance" }, { id: "leaves", label: "Leaves" }, { id: "payroll", label: "Payroll" }, { id: "recruitment", label: "Recruitment" }] },
  { id: "expenses", label: "Expenses" },
  { id: "fixedassets", label: "Fixed Assets" },
  { id: "manufacturing", label: "Manufacturing", subs: [{ id: "productionorders", label: "Production Orders" }, { id: "workcenters", label: "Work Centers" }, { id: "qualitycontrol", label: "Quality Control" }, { id: "maintenance", label: "Maintenance" }] },
  { id: "projects", label: "Projects" },
  { id: "service", label: "Service" },
  { id: "documents", label: "Documents" },
  { id: "reports", label: "Reports" },
  { id: "approvals", label: "Approvals" },
  { id: "users", label: "Users & Permissions" },
  { id: "settings", label: "Settings", subs: [{ id: "general", label: "General" }, { id: "currencies", label: "Currencies" }] },
];

const MODULE_ACCESS = {
  dashboard: ["Admin", "Factory Manager", "Sales Manager", "Accountant", "Viewer"],
  crm: ["Admin", "Sales Manager", "Viewer"],
  sales: ["Admin", "Sales Manager", "Viewer"],
  purchasing: ["Admin", "Factory Manager", "Viewer"],
  inventory: ["Admin", "Factory Manager", "Viewer"],
  storemgmt: ["Admin", "Factory Manager", "Viewer"],
  accounting: ["Admin", "Accountant", "Viewer"],
  hr: ["Admin", "Viewer"],
  expenses: ["Admin", "Accountant", "Viewer"],
  fixedassets: ["Admin", "Accountant", "Viewer"],
  manufacturing: ["Admin", "Factory Manager", "Viewer"],
  projects: ["Admin", "Viewer"],
  service: ["Admin", "Sales Manager", "Viewer"],
  documents: ["Admin", "Viewer"],
  reports: ["Admin", "Viewer", "Accountant", "Sales Manager", "Factory Manager"],
  approvals: ["Admin", "Viewer"],
  users: ["Admin"],
  settings: ["Admin"],
};

/* ---------------- Seed: master data ---------------- */

const seedSuppliers = [
  { id: "SUP-001", name: "Millhaven Textiles", contact: "orders@millhaven.co", lead: 5 },
  { id: "SUP-002", name: "PureBase Chemical Supply", contact: "sales@purebase.com", lead: 3 },
  { id: "SUP-003", name: "PackRight Industries", contact: "hello@packright.com", lead: 4 },
  { id: "SUP-004", name: "Guardian Electronics Ltd", contact: "sales@guardianelec.com", lead: 10 },
];

const seedMaterials = [
  { id: "RM-001", name: "Cotton Terry Fabric", unit: "m", stock: 420, reorder: 150, cost: 3.2, supplier: "SUP-001" },
  { id: "RM-002", name: "Bedsheet Cotton (300tc)", unit: "m", stock: 260, reorder: 200, cost: 4.1, supplier: "SUP-001" },
  { id: "RM-003", name: "Saponified Soap Base", unit: "kg", stock: 95, reorder: 100, cost: 2.6, supplier: "SUP-002" },
  { id: "RM-004", name: "Shampoo Base Liquid", unit: "L", stock: 180, reorder: 120, cost: 1.9, supplier: "SUP-002" },
  { id: "RM-005", name: "Fragrance Oil (Neutral)", unit: "L", stock: 22, reorder: 15, cost: 38.0, supplier: "SUP-002" },
  { id: "RM-006", name: "30ml Amenity Bottle", unit: "pc", stock: 3200, reorder: 1500, cost: 0.12, supplier: "SUP-003" },
  { id: "RM-007", name: "Kraft Retail Box (small)", unit: "pc", stock: 640, reorder: 800, cost: 0.35, supplier: "SUP-003" },
  { id: "RM-008", name: "Woven Label Tag", unit: "pc", stock: 5000, reorder: 2000, cost: 0.04, supplier: "SUP-003" },
  { id: "RM-009", name: "Electronic Lock Mechanism", unit: "pc", stock: 150, reorder: 100, cost: 12.50, supplier: "SUP-004" },
  { id: "RM-010", name: "Keypad Module", unit: "pc", stock: 200, reorder: 100, cost: 6.75, supplier: "SUP-004" },
  { id: "RM-011", name: "Safe Steel Body", unit: "pc", stock: 80, reorder: 50, cost: 45.0, supplier: "SUP-004" },
  { id: "RM-012", name: "Circuit Board Assembly", unit: "pc", stock: 120, reorder: 80, cost: 18.25, supplier: "SUP-004" },
];

const seedProducts = [
  { id: "FG-001", name: "Bath Towel (600gsm)", category: "Linens", unit: "pc", price: 14.5, reorder: 100,
    stockByBranch: { "BR-HQ": 140, "BR-DEN": 25, "BR-TPA": 15 }, bom: [{ id: "RM-001", qty: 0.9 }, { id: "RM-008", qty: 1 }] },
  { id: "FG-002", name: "Bed Sheet Set (Queen)", category: "Linens", unit: "set", price: 42.0, reorder: 50,
    stockByBranch: { "BR-HQ": 40, "BR-DEN": 10, "BR-TPA": 10 }, bom: [{ id: "RM-002", qty: 4.2 }, { id: "RM-008", qty: 1 }] },
  { id: "FG-003", name: "Shampoo 30ml Amenity", category: "Amenities", unit: "pc", price: 0.9, reorder: 1000,
    stockByBranch: { "BR-HQ": 2000, "BR-DEN": 250, "BR-TPA": 150 }, bom: [{ id: "RM-004", qty: 0.03 }, { id: "RM-005", qty: 0.002 }, { id: "RM-006", qty: 1 }] },
  { id: "FG-004", name: "Soap Bar 40g", category: "Amenities", unit: "pc", price: 0.55, reorder: 1200,
    stockByBranch: { "BR-HQ": 2600, "BR-DEN": 300, "BR-TPA": 200 }, bom: [{ id: "RM-003", qty: 0.04 }, { id: "RM-005", qty: 0.001 }, { id: "RM-007", qty: 0.2 }] },
  { id: "FG-005", name: "Guest Slippers (pair)", category: "Amenities", unit: "pair", price: 1.8, reorder: 300,
    stockByBranch: { "BR-HQ": 220, "BR-DEN": 40, "BR-TPA": 30 }, bom: [{ id: "RM-001", qty: 0.15 }, { id: "RM-007", qty: 1 }] },
  { id: "FG-006", name: "Digital Door Lock", category: "Devices", unit: "pc", price: 89.0, reorder: 40,
    stockByBranch: { "BR-HQ": 60, "BR-DEN": 8, "BR-TPA": 6 }, bom: [{ id: "RM-009", qty: 1 }, { id: "RM-010", qty: 1 }, { id: "RM-012", qty: 1 }] },
  { id: "FG-007", name: "In-Room Safe", category: "Devices", unit: "pc", price: 129.0, reorder: 25,
    stockByBranch: { "BR-HQ": 35, "BR-DEN": 5, "BR-TPA": 4 }, bom: [{ id: "RM-011", qty: 1 }, { id: "RM-012", qty: 1 }, { id: "RM-010", qty: 1 }] },
];

const seedCustomers = [
  { id: "CL-001", name: "Grand Meridian Hotel", location: "Chicago, IL", contact: "procurement@grandmeridian.com", branch: "BR-HQ" },
  { id: "CL-002", name: "Coastal Palms Resort", location: "Tampa, FL", contact: "supply@coastalpalms.com", branch: "BR-TPA" },
  { id: "CL-003", name: "Uptown Business Suites", location: "Denver, CO", contact: "ops@uptownsuites.com", branch: "BR-DEN" },
  { id: "CL-004", name: "Harbor View Inn", location: "Portland, ME", contact: "front.desk@harborview.com", branch: "BR-HQ" },
];

const seedLeads = [
  { id: "LD-0001", name: "Lakeside Convention Hotel", company: "Lakeside Group", source: "Trade Show", status: "Qualified", owner: "Jordan Blake", date: "2026-08-05" },
  { id: "LD-0002", name: "Marina Bay Suites", company: "Marina Hospitality", source: "Referral", status: "Contacted", owner: "Jordan Blake", date: "2026-08-09" },
  { id: "LD-0003", name: "Redwood Executive Inn", company: "Redwood Lodging", source: "Website", status: "New", owner: "Jordan Blake", date: "2026-08-15" },
];

const seedOpportunities = [
  { id: "OP-0001", customer: "CL-001", title: "2027 Linens Renewal Contract", value: 48000, stage: "Negotiation", closeDate: "2026-09-30" },
  { id: "OP-0002", customer: "CL-003", title: "New Property Amenity Rollout", value: 15500, stage: "Proposal", closeDate: "2026-10-15" },
  { id: "OP-0003", customer: "CL-002", title: "Soap Bar Formula Upgrade", value: 9200, stage: "Prospecting", closeDate: "2026-11-01" },
];

const seedQuotations = [
  { id: "QT-0001", client: "CL-004", date: "2026-08-13", items: [{ id: "FG-005", qty: 150 }, { id: "FG-003", qty: 300 }], status: "Accepted" },
  { id: "QT-0002", client: "CL-003", date: "2026-08-16", items: [{ id: "FG-002", qty: 20 }], status: "Sent" },
  { id: "QT-0003", client: "CL-002", date: "2026-08-17", items: [{ id: "FG-004", qty: 500 }], status: "Draft" },
];

const seedSalesOrders = [
  { id: "SO-001", client: "CL-001", branch: "BR-HQ", date: "2026-08-10", status: "Paid", items: [{ id: "FG-001", qty: 200, price: 14.5 }, { id: "FG-003", qty: 500, price: 0.9 }] },
  { id: "SO-002", client: "CL-002", branch: "BR-TPA", date: "2026-08-12", status: "Invoiced", items: [{ id: "FG-004", qty: 800, price: 0.55 }] },
  { id: "SO-003", client: "CL-003", branch: "BR-DEN", date: "2026-08-14", status: "Delivered", items: [{ id: "FG-002", qty: 40, price: 42.0 }] },
  { id: "SO-004", client: "CL-004", branch: "BR-HQ", date: "2026-08-16", status: "Pending", items: [{ id: "FG-005", qty: 150, price: 1.8 }, { id: "FG-003", qty: 300, price: 0.9 }] },
];

const seedReturns = [
  { id: "RT-0001", so: "SO-001", item: "FG-001", qty: 5, reason: "Damaged in transit", date: "2026-08-15", status: "Requested" },
];

const seedPurchaseRequests = [
  { id: "PR-0001", item: "RM-005", qty: 20, requestedBy: "Priya Nair", department: "Production", date: "2026-08-14", status: "Pending" },
  { id: "PR-0002", item: "RM-007", qty: 1500, requestedBy: "Priya Nair", department: "Production", date: "2026-08-15", status: "Converted" },
];

const seedPurchaseOrders = [
  { id: "PO-001", supplier: "SUP-002", date: "2026-08-11", status: "Received", items: [{ id: "RM-004", qty: 100, cost: 1.9 }, { id: "RM-005", qty: 10, cost: 38.0 }] },
  { id: "PO-002", supplier: "SUP-003", date: "2026-08-15", status: "Pending", items: [{ id: "RM-007", qty: 1000, cost: 0.35 }, { id: "RM-006", qty: 2000, cost: 0.12 }] },
];

const seedPurchaseInvoices = [
  { id: "PINV-0001", po: "PO-001", supplier: "SUP-002", amount: 570, date: "2026-08-11", status: "Paid" },
];

const seedProductionOrders = [
  { id: "PB-001", product: "FG-003", qty: 1000, date: "2026-08-09", status: "Completed" },
  { id: "PB-002", product: "FG-005", qty: 300, date: "2026-08-16", status: "Planned" },
];

const seedWorkCenters = [
  { id: "WC-001", name: "Injection Molding Line", department: "Devices", capacityPerDay: 200, status: "Running" },
  { id: "WC-002", name: "Electronics Assembly", department: "Devices", capacityPerDay: 150, status: "Running" },
  { id: "WC-003", name: "Textile Cutting & Sewing", department: "Linens", capacityPerDay: 500, status: "Running" },
  { id: "WC-004", name: "Amenity Filling & Packaging", department: "Amenities", capacityPerDay: 3000, status: "Idle" },
  { id: "WC-005", name: "Final Assembly & Testing", department: "Devices", capacityPerDay: 120, status: "Maintenance" },
];

const seedQualityChecks = [
  { id: "QC-0001", batch: "PB-001", product: "FG-003", inspector: "Aiko Tanaka", date: "2026-08-09", result: "Pass", notes: "Fill accuracy within tolerance." },
  { id: "QC-0002", batch: "PB-002", product: "FG-005", inspector: "Aiko Tanaka", date: "2026-08-16", result: "Pending", notes: "" },
];

const seedMaintenance = [
  { id: "MT-0001", equipment: "Injection Molding Press #1", workCenter: "WC-001", lastService: "2026-06-01", nextDue: "2026-09-01", status: "Up to Date", technician: "Diego Alvarez" },
  { id: "MT-0002", equipment: "Pick-and-Place Machine", workCenter: "WC-002", lastService: "2026-05-15", nextDue: "2026-08-15", status: "Due Soon", technician: "Diego Alvarez" },
  { id: "MT-0003", equipment: "Final Test Rig", workCenter: "WC-005", lastService: "2026-04-01", nextDue: "2026-07-01", status: "Overdue", technician: "Priya Nair" },
];

const seedTransfers = [
  { id: "TR-001", product: "FG-004", from: "BR-HQ", to: "BR-TPA", qty: 200, date: "2026-08-13" },
];

const seedMovements = [
  { id: "MV-0001", type: "IN", itemType: "Material", item: "RM-004", qty: 100, warehouse: "BR-HQ", date: "2026-08-11", ref: "PO-001" },
  { id: "MV-0002", type: "IN", itemType: "Material", item: "RM-005", qty: 10, warehouse: "BR-HQ", date: "2026-08-11", ref: "PO-001" },
  { id: "MV-0003", type: "OUT", itemType: "Material", item: "RM-004", qty: 30, warehouse: "BR-HQ", date: "2026-08-09", ref: "PB-001" },
  { id: "MV-0004", type: "IN", itemType: "Product", item: "FG-003", qty: 1000, warehouse: "BR-HQ", date: "2026-08-09", ref: "PB-001" },
  { id: "MV-0005", type: "OUT", itemType: "Product", item: "FG-001", qty: 200, warehouse: "BR-HQ", date: "2026-08-10", ref: "SO-001" },
  { id: "MV-0006", type: "OUT", itemType: "Product", item: "FG-003", qty: 500, warehouse: "BR-HQ", date: "2026-08-10", ref: "SO-001" },
];

const seedStockCounts = [
  { id: "SC-0001", warehouse: "BR-HQ", item: "FG-001", systemQty: 140, countedQty: 138, date: "2026-08-15", status: "Completed" },
];

const seedItemLimits = [
  { id: "LIM-0001", item: "FG-001", warehouse: "BR-HQ", min: 80, max: 250 },
  { id: "LIM-0002", item: "FG-001", warehouse: "BR-DEN", min: 15, max: 60 },
  { id: "LIM-0003", item: "FG-003", warehouse: "BR-HQ", min: 800, max: 3000 },
  { id: "LIM-0004", item: "RM-005", warehouse: "BR-HQ", min: 15, max: 50 },
];

const seedTaxRates = [
  { id: "TAX-0001", name: "Standard Sales Tax", rate: 7.25, type: "Sales Tax", appliesTo: "Sales", isActive: true },
  { id: "TAX-0002", name: "State VAT", rate: 5.0, type: "VAT", appliesTo: "Both", isActive: true },
  { id: "TAX-0003", name: "Import Duty", rate: 3.5, type: "Duty", appliesTo: "Purchases", isActive: false },
];

const seedCoA = [
  { code: "1000", name: "Cash", type: "Asset" },
  { code: "1010", name: "Bank - Operating", type: "Asset" },
  { code: "1200", name: "Accounts Receivable", type: "Asset" },
  { code: "1400", name: "Inventory", type: "Asset" },
  { code: "1600", name: "Fixed Assets", type: "Asset" },
  { code: "2000", name: "Accounts Payable", type: "Liability" },
  { code: "3000", name: "Owner's Equity", type: "Equity" },
  { code: "4000", name: "Sales Revenue", type: "Revenue" },
  { code: "5000", name: "Cost of Goods Sold", type: "Expense" },
  { code: "5100", name: "Payroll Expense", type: "Expense" },
  { code: "5200", name: "Operating Expenses", type: "Expense" },
];

const seedJournal = [
  { id: "JE-0001", date: "2026-08-10", memo: "Record sale to Grand Meridian", debit: "Accounts Receivable", credit: "Sales Revenue", amount: 3350 },
  { id: "JE-0002", date: "2026-08-11", memo: "Purchase of shampoo base & fragrance oil", debit: "Inventory", credit: "Accounts Payable", amount: 570 },
];

const seedBanks = [
  { id: "BANK-01", name: "First Commerce Bank — Operating", number: "****4821", balance: 48250 },
  { id: "BANK-02", name: "First Commerce Bank — Payroll", number: "****9013", balance: 12500 },
];

const seedEmployees = [
  { id: "EMP-001", name: "Priya Nair", role: "Factory Manager", department: "Production", branch: "BR-HQ", email: "priya.nair@codepulse.co", hired: "2022-03-01", status: "Active" },
  { id: "EMP-002", name: "Jordan Blake", role: "Sales Manager", department: "Sales", branch: "BR-HQ", email: "jordan.blake@codepulse.co", hired: "2022-07-15", status: "Active" },
  { id: "EMP-003", name: "Casey Wu", role: "Accountant", department: "Finance", branch: "BR-HQ", email: "casey.wu@codepulse.co", hired: "2023-01-10", status: "Active" },
  { id: "EMP-004", name: "Morgan Lee", role: "Warehouse Associate", department: "Inventory", branch: "BR-DEN", email: "morgan.lee@codepulse.co", hired: "2023-09-05", status: "Active" },
  { id: "EMP-005", name: "Diego Alvarez", role: "Line Operator", department: "Production", branch: "BR-HQ", email: "diego.alvarez@codepulse.co", hired: "2024-02-20", status: "Active" },
  { id: "EMP-006", name: "Aiko Tanaka", role: "QC Inspector", department: "Production", branch: "BR-HQ", email: "aiko.tanaka@codepulse.co", hired: "2024-05-11", status: "Active" },
];

const seedAttendance = [
  { id: "AT-0001", employee: "Priya Nair", date: "2026-08-17", status: "Present" },
  { id: "AT-0002", employee: "Diego Alvarez", date: "2026-08-17", status: "Present" },
  { id: "AT-0003", employee: "Aiko Tanaka", date: "2026-08-17", status: "Late" },
  { id: "AT-0004", employee: "Morgan Lee", date: "2026-08-17", status: "Present" },
];

const seedLeaves = [
  { id: "LV-0001", employee: "Diego Alvarez", type: "Sick", from: "2026-08-18", to: "2026-08-19", status: "Pending" },
  { id: "LV-0002", employee: "Aiko Tanaka", type: "Vacation", from: "2026-08-25", to: "2026-08-29", status: "Approved" },
];

const seedPayroll = [
  { id: "PR-2026-08-EMP001", employee: "Priya Nair", month: "2026-08", base: 6200, deductions: 450, status: "Paid" },
  { id: "PR-2026-08-EMP002", employee: "Jordan Blake", month: "2026-08", base: 5400, deductions: 380, status: "Paid" },
  { id: "PR-2026-08-EMP003", employee: "Casey Wu", month: "2026-08", base: 5100, deductions: 360, status: "Paid" },
  { id: "PR-2026-08-EMP005", employee: "Diego Alvarez", month: "2026-08", base: 3600, deductions: 240, status: "Draft" },
];

const seedRecruitment = [
  { id: "REC-0001", position: "Line Operator", department: "Production", candidate: "Sam Ortiz", stage: "Interview" },
  { id: "REC-0002", position: "Warehouse Associate", department: "Inventory", candidate: "Lena Fischer", stage: "Applied" },
  { id: "REC-0003", position: "Sales Coordinator", department: "Sales", candidate: "Marco Silva", stage: "Offer" },
];

const seedExpenses = [
  { id: "EXP-0001", employee: "Diego Alvarez", category: "Travel", amount: 120, date: "2026-08-12", status: "Approved" },
  { id: "EXP-0002", employee: "Jordan Blake", category: "Client Entertainment", amount: 340, date: "2026-08-13", status: "Pending" },
  { id: "EXP-0003", employee: "Priya Nair", category: "Equipment Repair", amount: 560, date: "2026-08-14", status: "Paid" },
];

const seedFixedAssets = [
  { id: "FA-0001", name: "Industrial Sewing Line", category: "Production Equipment", purchaseDate: "2023-01-15", cost: 85000, lifeYears: 10 },
  { id: "FA-0002", name: "Delivery Truck (Ford Transit)", category: "Vehicle", purchaseDate: "2022-06-01", cost: 42000, lifeYears: 7 },
  { id: "FA-0003", name: "Warehouse Forklift", category: "Equipment", purchaseDate: "2024-02-10", cost: 18000, lifeYears: 8 },
  { id: "FA-0004", name: "Office Computers (x8)", category: "IT Equipment", purchaseDate: "2025-05-20", cost: 9600, lifeYears: 4 },
];

const seedProjects = [
  { id: "PRJ-0001", name: "New Amenity Line Setup", client: "Internal", manager: "Priya Nair", status: "In Progress", budget: 25000, start: "2026-07-01", end: "2026-09-30", percent: 60 },
  { id: "PRJ-0002", name: "Grand Meridian Rebrand Rollout", client: "Grand Meridian Hotel", manager: "Jordan Blake", status: "Planning", budget: 8000, start: "2026-09-01", end: "2026-10-15", percent: 10 },
  { id: "PRJ-0003", name: "Warehouse Automation Pilot", client: "Internal", manager: "Morgan Lee", status: "Completed", budget: 15000, start: "2026-03-01", end: "2026-06-30", percent: 100 },
];

const seedService = [
  { id: "SV-0001", client: "CL-002", issue: "Laundry bag zipper defects on last batch", priority: "High", status: "Open", assignedTo: "Aiko Tanaka", date: "2026-08-15" },
  { id: "SV-0002", client: "CL-001", issue: "Reorder amenity tray holders", priority: "Low", status: "Resolved", assignedTo: "Jordan Blake", date: "2026-08-11" },
  { id: "SV-0003", client: "CL-003", issue: "Damaged bed sheet set in latest delivery", priority: "Medium", status: "In Progress", assignedTo: "Priya Nair", date: "2026-08-14" },
];

const seedDocuments = [
  { id: "DOC-0001", name: "Supplier Agreement - Millhaven Textiles.pdf", module: "Purchasing", type: "Contract", uploadedBy: "Sam Rivera", date: "2026-02-10" },
  { id: "DOC-0002", name: "Q2 Financial Statement.xlsx", module: "Accounting", type: "Report", uploadedBy: "Casey Wu", date: "2026-07-05" },
  { id: "DOC-0003", name: "Fire Safety Certificate.pdf", module: "HR", type: "Compliance", uploadedBy: "Priya Nair", date: "2026-01-20" },
  { id: "DOC-0004", name: "Grand Meridian Hotel MSA.pdf", module: "Sales", type: "Contract", uploadedBy: "Jordan Blake", date: "2025-11-01" },
];

/* ---------------- UI atoms ---------------- */

const STATUS_COLOR = {
  Pending: "#8A6A2E", Fulfilled: "#3A5C86", Delivered: "#3A5C86", Invoiced: "#6B4FA0", Paid: "#3F7D5C",
  Planned: "#8A6A2E", Completed: "#3F7D5C", Received: "#3F7D5C", Draft: "var(--text-label)", Sent: "#3A5C86",
  Accepted: "#3F7D5C", Rejected: "#A64B3A", Converted: "#3F7D5C", Approved: "#3F7D5C", Requested: "#8A6A2E",
  Restocked: "#3F7D5C", New: "#8A6A2E", Contacted: "#3A5C86", Qualified: "#3F7D5C", Lost: "#A64B3A",
  Prospecting: "#8A6A2E", Proposal: "#3A5C86", Negotiation: "#6B4FA0", Won: "#3F7D5C", Unpaid: "#8A6A2E",
  Open: "#8A6A2E", "In Progress": "#3A5C86", Resolved: "#3F7D5C", Active: "#3F7D5C",
  Applied: "#8A6A2E", Interview: "#3A5C86", Offer: "#6B4FA0", Hired: "#3F7D5C",
  Planning: "#8A6A2E", "On Hold": "#A64B3A", Present: "#3F7D5C", Absent: "#A64B3A", Late: "#8A6A2E",
  "Raw Material": "#3A5C86", "Finished Good": "#3F7D5C", Supplier: "#3A5C86", Expense: "#8A6A2E",
  "Purchase Request": "#8A6A2E", "Leave Request": "#3A5C86", "Expense Claim": "#6B4FA0",
  "Below Min": "#A64B3A", "Within Range": "#3F7D5C", "Above Max": "#6B4FA0",
  Running: "#3F7D5C", Idle: "#8A6A2E", Maintenance: "#A64B3A", Fail: "#A64B3A", Pass: "#3F7D5C",
  "Up to Date": "#3F7D5C", "Due Soon": "#8A6A2E", Overdue: "#A64B3A",
  "Sales Order": "#3A5C86", "Manual": "#8A6A2E",
};

function Pill({ children }) {
  const { t } = useApp();
  const c = STATUS_COLOR[children] || "var(--text-secondary)";
  return (
    <span className="inline-flex items-center px-2 py-0.5 rounded-sm text-[11px] font-semibold tracking-wide uppercase"
      style={{ color: c, background: `${c}1A`, border: `1px solid ${c}44` }}>{t(children)}</span>
  );
}
function RolePill({ role }) {
  const { t } = useApp();
  return (
    <span className="inline-flex items-center px-2 py-0.5 rounded-sm text-[10px] font-semibold tracking-wide uppercase"
      style={{ color: "#C08A2E", background: "#C08A2E1A", border: "1px solid #C08A2E44" }}>{t(role)}</span>
  );
}
function Gauge({ value, max, label }) {
  const pct = Math.min(100, Math.round((value / max) * 100));
  const low = value < max * 0.35;
  return (
    <div className="w-full">
      <div className="flex justify-between text-[11px] mb-1" style={{ color: "var(--text-secondary)" }}><span>{label}</span><span className="font-mono">{value}</span></div>
      <div className="h-1.5 w-full rounded-full" style={{ background: "var(--border)" }}>
        <div className="h-1.5 rounded-full" style={{ width: `${pct}%`, background: low ? "#A64B3A" : "#C08A2E" }} />
      </div>
    </div>
  );
}
function Card({ title, value, sub, accent }) {
  return (
    <div className="p-4 rounded-md" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
      <div className="text-[11px] uppercase tracking-widest font-semibold" style={{ color: "var(--text-label)" }}>{title}</div>
      <div className="text-2xl font-bold mt-1" style={{ color: accent || "var(--heading)", fontFamily: "'Oswald', sans-serif" }}>{value}</div>
      {sub && <div className="text-xs mt-1" style={{ color: "var(--text-secondary)" }}>{sub}</div>}
    </div>
  );
}
// Exports the actual rendered <table> DOM element — whatever the user is
// currently looking at (translated text, formatted numbers, everything) —
// straight to Excel or CSV. Because it reads the real DOM instead of needing
// a hand-built data array per page, every table in the app gets export for
// free just by using this shared component.
function exportTableElement(tableEl, name, format, reportTitle) {
  if (!tableEl) return;
  const safeName = (name || "CodePulse_Export").replace(/[^\w\-]+/g, "_").slice(0, 60);
  const rawWs = XLSX.utils.table_to_sheet(tableEl);
  if (!rawWs["!ref"]) { alert("Nothing to export yet."); return; }

  // Rebuild the sheet with a company/title/date header block above the real
  // data, then reapply column widths and autofilter against the shifted
  // range. (Verified against what the free SheetJS build actually writes:
  // merged cells, row heights, column widths, and autofilter all persist in
  // the real .xlsx file. Font size/bold/fill color, and embedded images, do
  // NOT — those are paid SheetJS Pro features, confirmed by testing rather
  // than assumed. The HTML export option below covers full styling + logo.)
  const aoa = XLSX.utils.sheet_to_json(rawWs, { header: 1, raw: true });
  const colCount = aoa.reduce((max, row) => Math.max(max, row.length), 1);
  const headerRows = [
    ["CodePulse"],
    [reportTitle || safeName.replace(/_/g, " ")],
    [`Generated: ${todayStr()}`],
    [],
  ];
  const fullAoa = [...headerRows, ...aoa];
  const ws = XLSX.utils.aoa_to_sheet(fullAoa);
  const HEADER_ROWS = headerRows.length; // rows consumed before the real header/data

  ws["!merges"] = [0, 1, 2].map((r) => ({ s: { r, c: 0 }, e: { r, c: colCount - 1 } }));
  ws["!rows"] = [{ hpt: 24 }, { hpt: 18 }, { hpt: 14 }, { hpt: 6 }];

  const dataRange = XLSX.utils.decode_range(ws["!ref"]);
  const cols = [];
  for (let C = dataRange.s.c; C <= dataRange.e.c; C++) {
    let maxLen = 8;
    for (let R = HEADER_ROWS; R <= dataRange.e.r; R++) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      if (cell && cell.v != null) maxLen = Math.max(maxLen, String(cell.v).length);
    }
    cols.push({ wch: Math.min(maxLen + 2, 42) });
  }
  ws["!cols"] = cols;
  ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: HEADER_ROWS, c: dataRange.s.c }, e: { r: dataRange.e.r, c: dataRange.e.c } }) };
  ws["!margins"] = { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 };

  if (format === "csv") {
    const csv = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `${safeName}.csv`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } else {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, safeName.slice(0, 31));
    XLSX.writeFile(wb, `${safeName}.xlsx`);
  }
}

// Full-design export with the real logo and unlimited CSS styling — the
// route for "print or email this with our branding," since true .xlsx files
// from the free SheetJS library can't embed images or apply bold/color
// (confirmed above). This is a real, standalone HTML file: opens correctly
// in any browser or in Excel/Word, can be printed with your OS's normal
// print dialog once it's outside this preview, and can be attached to or
// pasted into an email.
function downloadHTML(tableEl, name, reportTitle) {
  if (!tableEl) return;
  const safeName = (name || "CodePulse_Export").replace(/[^\w\-]+/g, "_").slice(0, 60);
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${reportTitle || safeName}</title>
  <style>
    body{font-family:Arial,Helvetica,sans-serif;color:#23271F;padding:32px;max-width:1000px;margin:0 auto}
    .header{display:flex;align-items:center;gap:16px;border-bottom:3px solid #C08A2E;padding-bottom:16px;margin-bottom:22px}
    .header img{height:52px;width:auto;border-radius:4px}
    .company{font-size:22px;font-weight:700;color:#1B2421;letter-spacing:0.5px}
    .report-title{font-size:16px;color:#3A4038;margin-top:2px}
    .date{font-size:12px;color:#8A8578;margin-top:3px}
    table{width:100%;border-collapse:collapse;font-size:13px;margin-top:8px}
    th{background:#1B2421;color:#fff;padding:9px 10px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:0.4px}
    td{padding:8px 10px;border-bottom:1px solid #E4E0D4}
    tr:nth-child(even) td{background:#F7F5F1}
    footer{margin-top:24px;font-size:11px;color:#8A8578;border-top:1px solid #E4E0D4;padding-top:10px}
    @media print { body{padding:16px} }
  </style></head>
  <body>
    <div class="header">
      <img src="${LOGO_DATA_URI}" alt="CodePulse" />
      <div>
        <div class="company">CODEPULSE</div>
        <div class="report-title">${reportTitle || safeName.replace(/_/g, " ")}</div>
        <div class="date">Generated: ${todayStr()}</div>
      </div>
    </div>
    ${tableEl.outerHTML}
    <footer>CodePulse Enterprise Console — exported ${todayStr()}</footer>
  </body></html>`;
  const blob = new Blob([html], { type: "text/html;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = `${safeName}.html`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function Table({ columns, rows, renderRow, title }) {
  const { t } = useApp();
  const tableRef = useRef(null);
  const [open, setOpen] = useState(false);
  const exportName = `${title || "CodePulse_Export"}_${todayStr()}`;

  return (
    <div>
      {rows.length > 0 && (
        <div className="flex justify-end mb-1.5">
          <div className="relative inline-block">
            <Button variant="ghost" onClick={() => setOpen((o) => !o)}>Export ▾</Button>
            {open && (
              <>
                <div className="fixed inset-0 z-10" onClick={() => setOpen(false)} />
                <div className="absolute right-0 z-20 mt-1 rounded-sm overflow-hidden" style={{ background: "var(--surface)", border: "1px solid var(--border)", minWidth: 140, boxShadow: "0 4px 14px rgba(0,0,0,0.12)" }}>
                  {[
                    { label: "Excel (.xlsx)", action: () => exportTableElement(tableRef.current, exportName, "xlsx", title) },
                    { label: "CSV", action: () => exportTableElement(tableRef.current, exportName, "csv", title) },
                    { label: "HTML (logo, for print/email)", action: () => downloadHTML(tableRef.current, exportName, title) },
                  ].map((opt) => (
                    <button key={opt.label} onClick={() => { setOpen(false); opt.action(); }}
                      className="w-full text-left px-3 py-2 text-xs" style={{ color: "var(--text)" }}
                      onMouseEnter={(e) => (e.currentTarget.style.background = "var(--surface-alt)")}
                      onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}>
                      {opt.label}
                    </button>
                  ))}
                </div>
              </>
            )}
          </div>
        </div>
      )}
      <div className="overflow-x-auto rounded-md" style={{ border: "1px solid var(--border)" }}>
        <table ref={tableRef} className="w-full text-sm" style={{ borderCollapse: "collapse" }}>
          <thead><tr style={{ background: "#1B2421" }}>
            {columns.map((c) => <th key={c} className="text-left px-3 py-2 text-[11px] uppercase tracking-wider font-semibold" style={{ color: "#E8E4D8" }}>{t(c)}</th>)}
          </tr></thead>
          <tbody>
            {rows.map((r, i) => <tr key={i} style={{ background: i % 2 ? "var(--surface-alt)" : "var(--surface)", borderTop: "1px solid var(--zebra-border)" }}>{renderRow(r)}</tr>)}
          </tbody>
        </table>
      </div>
    </div>
  );
}
function Td({ children, mono }) {
  return <td className="px-3 py-2 align-middle" style={{ fontFamily: mono ? "'IBM Plex Mono', monospace" : undefined, color: "var(--text)" }}>{children}</td>;
}
function Button({ children, onClick, variant = "primary", disabled }) {
  const { t } = useApp();
  const styles = {
    primary: { background: "#1B2421", color: "#F3EFE6", border: "1px solid #1B2421" },
    ghost: { background: "transparent", color: "var(--text)", border: "1px solid var(--border-strong)" },
    accent: { background: "#C08A2E", color: "#1B2421", border: "1px solid #C08A2E" },
    danger: { background: "transparent", color: "#A64B3A", border: "1px solid #A64B3A66" },
  };
  return (
    <button onClick={onClick} disabled={disabled}
      className="px-3 py-1.5 rounded-sm text-xs font-semibold uppercase tracking-wide transition-opacity mr-1.5"
      style={{ ...styles[variant], opacity: disabled ? 0.4 : 1, cursor: disabled ? "not-allowed" : "pointer" }}>{typeof children === "string" ? t(children) : children}</button>
  );
}
function SectionHeader({ title, sub }) {
  const { t } = useApp();
  return (
    <div className="mb-5">
      <h1 className="text-2xl font-bold" style={{ color: "var(--heading)", fontFamily: "'Oswald', sans-serif" }}>{t(title)}</h1>
      {sub && <p className="text-sm mt-0.5" style={{ color: "var(--text-secondary)" }}>{t(sub)}</p>}
    </div>
  );
}

/* Generic inline add-form driven by a field config array.
   Fields of type "money" render an amount input plus a currency picker;
   the entered amount is converted to USD (the system's storage currency) on submit. */
function AddForm({ fields, onSubmit, submitLabel = "+ Add", initialValues }) {
  const { currencies, currency: globalCurrency, toUSD, t } = useApp();
  const initial = Object.fromEntries(fields.map((f) => {
    const existing = initialValues?.[f.key];
    if (f.type === "money") {
      if (existing != null) {
        const rate = currencies.find((c) => c.code === globalCurrency)?.rate ?? 1;
        return [f.key, { amount: +(existing * rate).toFixed(2), currency: globalCurrency }];
      }
      return [f.key, { amount: f.default ?? 0, currency: globalCurrency }];
    }
    return [f.key, existing ?? f.default ?? (f.type === "number" ? 0 : (f.options ? (f.options[0].value ?? f.options[0]) : ""))];
  }));
  const [vals, setVals] = useState(initial);
  const set = (k, v) => setVals((old) => ({ ...old, [k]: v }));
  const setMoney = (k, patch) => setVals((old) => ({ ...old, [k]: { ...old[k], ...patch } }));

  function submit() {
    const out = { ...vals };
    fields.forEach((f) => { if (f.type === "money") out[f.key] = toUSD(vals[f.key].amount, vals[f.key].currency); });
    onSubmit(out);
  }

  return (
    <div className="p-4 rounded-md mb-5 flex items-end gap-3 flex-wrap" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
      {fields.map((f) => (
        <div key={f.key}>
          <div className="text-[11px] uppercase tracking-wide mb-1" style={{ color: "var(--text-label)" }}>{t(f.label)}</div>
          {f.type === "select" ? (
            <select value={vals[f.key]} onChange={(e) => set(f.key, e.target.value)} className="text-sm px-2 py-1.5 rounded-sm" style={{ border: "1px solid var(--border-strong)" }}>
              {f.options.map((o) => <option key={o.value ?? o} value={o.value ?? o}>{t(o.label ?? o)}</option>)}
            </select>
          ) : f.type === "money" ? (
            <div className="flex items-center gap-1">
              <input type="number" value={vals[f.key].amount} onChange={(e) => setMoney(f.key, { amount: +e.target.value })}
                className="text-sm px-2 py-1.5 rounded-sm" style={{ border: "1px solid var(--border-strong)", width: 90 }} />
              <select value={vals[f.key].currency} onChange={(e) => setMoney(f.key, { currency: e.target.value })}
                className="text-xs px-1 py-1.5 rounded-sm font-mono" style={{ border: "1px solid var(--border-strong)" }}>
                {currencies.map((c) => <option key={c.code} value={c.code}>{c.code}</option>)}
              </select>
            </div>
          ) : (
            <input type={f.type || "text"} value={vals[f.key]}
              onChange={(e) => set(f.key, f.type === "number" ? +e.target.value : e.target.value)}
              className="text-sm px-2 py-1.5 rounded-sm" style={{ border: "1px solid var(--border-strong)", width: f.type === "number" ? 90 : 170 }} />
          )}
        </div>
      ))}
      <Button variant="accent" onClick={submit}>{submitLabel}</Button>
    </div>
  );
}

function ListPage({ title, sub, addFields, onAdd, onUpdate, onDelete, canEdit, columns, rows, renderRow, emptyText, idKey = "id" }) {
  const { isAdmin } = useApp();
  const [editingId, setEditingId] = useState(null);
  const editingRow = rows.find((r) => r[idKey] === editingId);
  const hasActions = !!(onUpdate || onDelete);
  const finalColumns = hasActions ? [...columns, ""] : columns;
  function finalRenderRow(r) {
    return (
      <>
        {renderRow(r)}
        {hasActions && (
          <Td>
            {isAdmin && onUpdate && <Button variant="ghost" onClick={() => setEditingId(r[idKey])}>Edit</Button>}
            {canEdit !== false && onDelete && <Button variant="danger" onClick={() => onDelete(r[idKey])}>Delete</Button>}
          </Td>
        )}
      </>
    );
  }
  return (
    <div>
      <SectionHeader title={title} sub={sub} />
      {editingRow ? (
        <AddForm key={editingId} fields={addFields} initialValues={editingRow}
          onSubmit={(vals) => { onUpdate(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      ) : (canEdit !== false && addFields && <AddForm fields={addFields} onSubmit={onAdd} />)}
      <Table title={title} columns={finalColumns} rows={rows} renderRow={finalRenderRow} />
      {rows.length === 0 && emptyText && <div className="text-sm mt-3" style={{ color: "var(--text-secondary)" }}>{emptyText}</div>}
    </div>
  );
}

/* ---------------- App Context ---------------- */

const Ctx = createContext(null);
const useApp = () => useContext(Ctx);

/* ---------------- Login screen ---------------- */

function Login({ onLogin, lang, setLang, t, users }) {
  const [selected, setSelected] = useState(users[0].id);
  const [password, setPassword] = useState("");
  const dir = lang === "ar" ? "rtl" : "ltr";
  return (
    <div dir={dir} className="w-full min-h-[700px] flex items-center justify-center" style={{ background: "#1B2421", fontFamily: "'Inter', system-ui, sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');`}</style>
      <div className="w-[380px] p-7 rounded-md" style={{ background: "#F3EFE6" }}>
        <div className="flex items-center justify-between mb-1">
          <div className="flex items-center gap-2.5">
            <img src={LOGO_DATA_URI} alt="CodePulse" style={{ height: 34, width: "auto", borderRadius: 4 }} />
            <div className="text-xl font-bold tracking-wide" style={{ color: "#1B2421", fontFamily: "'Oswald', sans-serif" }}>CODEPULSE</div>
          </div>
          <select value={lang} onChange={(e) => setLang(e.target.value)} className="text-xs px-2 py-1 rounded-sm" style={{ border: "1px solid #C7C2B2" }}>
            {LANGUAGES.map((l) => <option key={l.code} value={l.code}>{l.label}</option>)}
          </select>
        </div>
        <div className="text-[11px] uppercase tracking-widest mb-6" style={{ color: "#C08A2E" }}>{t("Enterprise Console — Sign In")}</div>
        <div className="text-[11px] uppercase tracking-wide mb-1.5" style={{ color: "#8A8578" }}>{t("Account")}</div>
        <div className="space-y-1.5 mb-4">
          {users.map((u) => (
            <button key={u.id} onClick={() => setSelected(u.id)}
              className="w-full flex items-center justify-between px-3 py-2 rounded-sm text-left text-sm"
              style={{ background: selected === u.id ? "#1B2421" : "#FFFFFF", color: selected === u.id ? "#F3EFE6" : "#23271F", border: "1px solid #E4E0D4" }}>
              <span>{u.name}</span>
              <span className="text-[10px] font-mono uppercase" style={{ color: selected === u.id ? "#C08A2E" : "#8A8578" }}>{t(u.role)}</span>
            </button>
          ))}
        </div>
        <div className="text-[11px] uppercase tracking-wide mb-1.5" style={{ color: "#8A8578" }}>{t("Password (demo — any value works)")}</div>
        <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••"
          className="w-full text-sm px-3 py-2 rounded-sm mb-5" style={{ border: "1px solid #C7C2B2", background: "#FFFFFF" }} />
        <button onClick={() => onLogin(users.find((u) => u.id === selected))}
          className="px-3 py-1.5 rounded-sm text-xs font-semibold uppercase tracking-wide"
          style={{ background: "#C08A2E", color: "#1B2421", border: "1px solid #C08A2E" }}>{t("Sign In")}</button>
        <div className="text-[11px] mt-4" style={{ color: "#8A8578" }}>{t("Modules and actions adapt to the signed-in role — Viewer is read-only everywhere.")}</div>
      </div>
    </div>
  );
}

/* ================================================================
   Root App
   ================================================================ */

export default function App() {
  const [user, setUser] = useState(null);
  const [selectedModule, setSelectedModule] = useState("dashboard");
  const [selectedSub, setSelectedSub] = useState(null);
  const [currency, setCurrency] = useState("USD");
  const [lang, setLang] = useState("en");
  const [theme, setTheme] = useState("light");

  // Every collection below is now backed by Firestore via the shared
  // useFirestoreCollection hook: it seeds itself from the app's existing
  // seed data the first time the collection is empty, then stays live via
  // Firestore's realtime listener. Session-only UI state (selected module,
  // language, theme, the signed-in user) intentionally stays local — there's
  // no reason someone else's currency display preference should sync to you.
  const usersCol = useFirestoreCollection("users", seedUsers);
  const materialsCol = useFirestoreCollection("materials", seedMaterials);
  const productsCol = useFirestoreCollection("products", seedProducts);
  const customersCol = useFirestoreCollection("customers", seedCustomers);
  const suppliersCol = useFirestoreCollection("suppliers", seedSuppliers);
  const leadsCol = useFirestoreCollection("leads", seedLeads);
  const opportunitiesCol = useFirestoreCollection("opportunities", seedOpportunities);
  const quotationsCol = useFirestoreCollection("quotations", seedQuotations);
  const salesOrdersCol = useFirestoreCollection("salesOrders", seedSalesOrders);
  const returnsCol = useFirestoreCollection("returns", seedReturns);
  const purchaseRequestsCol = useFirestoreCollection("purchaseRequests", seedPurchaseRequests);
  const purchaseOrdersCol = useFirestoreCollection("purchaseOrders", seedPurchaseOrders);
  const purchaseInvoicesCol = useFirestoreCollection("purchaseInvoices", seedPurchaseInvoices);
  const productionOrdersCol = useFirestoreCollection("productionOrders", seedProductionOrders);
  const workCentersCol = useFirestoreCollection("workCenters", seedWorkCenters);
  const qualityChecksCol = useFirestoreCollection("qualityChecks", seedQualityChecks);
  const maintenanceCol = useFirestoreCollection("maintenance", seedMaintenance);
  const transfersCol = useFirestoreCollection("transfers", seedTransfers);
  const movementsCol = useFirestoreCollection("movements", seedMovements);
  const stockCountsCol = useFirestoreCollection("stockCounts", seedStockCounts);
  const itemLimitsCol = useFirestoreCollection("itemLimits", seedItemLimits);
  const currenciesCol = useFirestoreCollection("currencies", seedCurrencies);
  const coaCol = useFirestoreCollection("coa", seedCoA);
  const taxRatesCol = useFirestoreCollection("taxRates", seedTaxRates);
  const journalCol = useFirestoreCollection("journal", seedJournal);
  const banksCol = useFirestoreCollection("banks", seedBanks);
  const employeesCol = useFirestoreCollection("employees", seedEmployees);
  const attendanceCol = useFirestoreCollection("attendance", seedAttendance);
  const leavesCol = useFirestoreCollection("leaves", seedLeaves);
  const payrollCol = useFirestoreCollection("payroll", seedPayroll);
  const recruitmentCol = useFirestoreCollection("recruitment", seedRecruitment);
  const expensesCol = useFirestoreCollection("expenses", seedExpenses);
  const fixedAssetsCol = useFirestoreCollection("fixedAssets", seedFixedAssets);
  const projectsCol = useFirestoreCollection("projects", seedProjects);
  const serviceCol = useFirestoreCollection("service", seedService);
  const documentsCol = useFirestoreCollection("documents", seedDocuments);
  const manualReceiptsCol = useFirestoreCollection("manualReceipts", []);
  const manualPaymentsCol = useFirestoreCollection("manualPayments", []);

  const users = usersCol.data, materials = materialsCol.data, products = productsCol.data;
  const customers = customersCol.data, suppliers = suppliersCol.data, leads = leadsCol.data;
  const opportunities = opportunitiesCol.data, quotations = quotationsCol.data, salesOrders = salesOrdersCol.data;
  const returns = returnsCol.data, purchaseRequests = purchaseRequestsCol.data, purchaseOrders = purchaseOrdersCol.data;
  const purchaseInvoices = purchaseInvoicesCol.data, productionOrders = productionOrdersCol.data;
  const workCenters = workCentersCol.data, qualityChecks = qualityChecksCol.data, maintenance = maintenanceCol.data;
  const transfers = transfersCol.data, movements = movementsCol.data, stockCounts = stockCountsCol.data;
  const itemLimits = itemLimitsCol.data, currencies = currenciesCol.data, coa = coaCol.data, taxRates = taxRatesCol.data;
  const journal = journalCol.data, banks = banksCol.data, employees = employeesCol.data, attendance = attendanceCol.data;
  const leaves = leavesCol.data, payroll = payrollCol.data, recruitment = recruitmentCol.data, expenses = expensesCol.data;
  const fixedAssets = fixedAssetsCol.data, projects = projectsCol.data, service = serviceCol.data, documents = documentsCol.data;
  const manualReceipts = manualReceiptsCol.data, manualPayments = manualPaymentsCol.data;

  // Firestore's realtime listeners populate these collections asynchronously
  // on first load — until the essentials arrive, render a loading screen
  // rather than a broken dashboard referencing data that doesn't exist yet.
  const stillLoading = usersCol.loading || materialsCol.loading || productsCol.loading || customersCol.loading || currenciesCol.loading;
  const firstError = [usersCol, materialsCol, productsCol].map((c) => c.error).find(Boolean);

  const matById = useMemo(() => Object.fromEntries(materials.map((m) => [m.id, m])), [materials]);
  const prodById = useMemo(() => Object.fromEntries(products.map((p) => [p.id, p])), [products]);
  const clientById = useMemo(() => Object.fromEntries(customers.map((c) => [c.id, c])), [customers]);
  const supById = useMemo(() => Object.fromEntries(suppliers.map((s) => [s.id, s])), [suppliers]);
  const branchById = useMemo(() => Object.fromEntries(BRANCHES.map((b) => [b.id, b])), []);

  const rateMap = useMemo(() => Object.fromEntries(currencies.map((c) => [c.code, c.rate])), [currencies]);
  const money = (amt) => formatMoney(amt, currency, rateMap);
  const toUSD = (amt, code) => amt / (rateMap[code] ?? 1);
  const t = (str) => (typeof str === "string" && I18N[lang]?.[str]) || str;
  const dir = lang === "ar" ? "rtl" : "ltr";
  const totalStock = (p) => Object.values(p.stockByBranch).reduce((s, v) => s + v, 0);
  const orderTotal = (o) => o.items.reduce((s, it) => s + it.qty * it.price, 0);
  const poTotal = (o) => o.items.reduce((s, it) => s + it.qty * it.cost, 0);

  const role = user?.role;
  const canEdit = !!role && role !== "Viewer";
  const isAdmin = role === "Admin";

  function reportError(e) { alert(e.message || String(e)); }

  /* ---- Actions: Purchasing ---- */
  function addPR(vals) { purchaseRequestsCol.add({ ...vals, date: todayStr(), status: "Pending" }, uid("PR")); }
  function updatePR(id, vals) { purchaseRequestsCol.update(id, vals); }
  function approvePR(id, decision) { purchaseRequestsCol.update(id, { status: decision }); }
  function convertPRtoPO(pr) {
    const material = matById[pr.item];
    newPO(material.supplier, [{ id: pr.item, qty: pr.qty, cost: material.cost }]);
    purchaseRequestsCol.update(pr.id, { status: "Converted" });
  }
  function newPO(supplierId, items) {
    purchaseOrdersCol.add({ supplier: supplierId, date: todayStr(), status: "Pending", items }, uid("PO"));
  }
  function receivePO(id) {
    const po = purchaseOrders.find((p) => p.id === id);
    receivePurchaseOrder(po).catch(reportError);
  }
  function payPurchaseInvoice(id) { purchaseInvoicesCol.update(id, { status: "Paid" }); }

  /* ---- Actions: Manufacturing ---- */
  function canRunBatch(product, qty) { return product.bom.every((b) => matById[b.id].stock >= b.qty * qty); }
  function completeBatch(id) {
    const batch = productionOrders.find((p) => p.id === id);
    const product = prodById[batch.product];
    completeProductionBatch(batch, product, matById).catch(reportError);
  }
  function newBatch(productId, qty) { productionOrdersCol.add({ product: productId, qty, date: todayStr(), status: "Planned" }, uid("PB")); }

  /* ---- Actions: Manufacturing operations (Work Centers / Quality Control / Maintenance) ---- */
  function addWorkCenter(vals) { workCentersCol.add({ ...vals, status: vals.status || "Idle" }, uid("WC")); }
  function updateWorkCenter(id, vals) { workCentersCol.update(id, vals); }
  function deleteWorkCenter(id) { workCentersCol.remove(id); }
  function addQualityCheck(vals) { qualityChecksCol.add({ ...vals, date: todayStr() }, uid("QC")); }
  function updateQualityCheck(id, vals) { qualityChecksCol.update(id, vals); }
  function deleteQualityCheck(id) { qualityChecksCol.remove(id); }
  function addMaintenance(vals) { maintenanceCol.add(vals, uid("MT")); }
  function updateMaintenance(id, vals) { maintenanceCol.update(id, vals); }
  function deleteMaintenance(id) { maintenanceCol.remove(id); }
  function logMaintenanceService(id) {
    const today = todayStr();
    const next = new Date(today); next.setMonth(next.getMonth() + 3);
    maintenanceCol.update(id, { lastService: today, nextDue: next.toISOString().slice(0, 10), status: "Up to Date" });
  }

  /* ---- Actions: Inventory ---- */
  function transferStock(productId, from, to, qty) {
    transferProductStock(productId, from, to, qty).catch(reportError);
  }
  function addStockCount(vals) {
    const isMaterial = vals.itemId.startsWith("RM");
    const systemQty = isMaterial ? (vals.warehouse === "BR-HQ" ? matById[vals.itemId].stock : 0) : prodById[vals.itemId].stockByBranch[vals.warehouse];
    stockCountsCol.add({ warehouse: vals.warehouse, item: vals.itemId, systemQty, countedQty: vals.countedQty, date: todayStr(), status: "Completed" }, uid("SC"));
  }

  /* ---- Actions: Store Management ---- */
  function addStoreItem(vals) {
    const isMaterial = vals.itemId.startsWith("RM");
    if (isMaterial && vals.warehouse !== "BR-HQ") { alert("Raw materials are only stocked at the factory warehouse."); return; }
    addStoreItemTx({ itemId: vals.itemId, warehouse: vals.warehouse, qty: vals.qty, reason: vals.reason, isMaterial }).catch(reportError);
  }
  function removeStoreItem(vals) {
    const isMaterial = vals.itemId.startsWith("RM");
    if (isMaterial && vals.warehouse !== "BR-HQ") { alert("Raw materials are only stocked at the factory warehouse."); return; }
    const itemName = isMaterial ? matById[vals.itemId]?.name : prodById[vals.itemId]?.name;
    removeStoreItemTx({ itemId: vals.itemId, warehouse: vals.warehouse, qty: vals.qty, reason: vals.reason, isMaterial, itemName }).catch(reportError);
  }
  function upsertLimit(vals) {
    const existing = itemLimits.find((l) => l.item === vals.item && l.warehouse === vals.warehouse);
    if (existing) itemLimitsCol.update(existing.id, { min: vals.min, max: vals.max });
    else itemLimitsCol.add({ item: vals.item, warehouse: vals.warehouse, min: vals.min, max: vals.max }, uid("LIM"));
  }
  function deleteItemLimit(id) { itemLimitsCol.remove(id); }

  /* ---- Actions: Currencies ---- */
  function addCurrency(vals) {
    if (currencies.some((c) => c.code === vals.code.toUpperCase())) { alert("That currency code already exists."); return; }
    currenciesCol.add({ code: vals.code.toUpperCase(), rate: vals.rate }, vals.code.toUpperCase());
  }
  function updateCurrencyRate(code, rate) { currenciesCol.update(code, { rate }); }

  /* ---- Actions: CRM / Sales ---- */
  function updateCustomer(id, vals) { customersCol.update(id, vals); }
  function updateEmployee(id, vals) { employeesCol.update(id, vals); }
  function addLead(vals) { leadsCol.add({ ...vals, status: "New", date: todayStr() }, uid("LD")); }
  function updateLead(id, vals) { leadsCol.update(id, vals); }
  function deleteLead(id) { leadsCol.remove(id); }
  function addOpportunity(vals) { opportunitiesCol.add({ ...vals, stage: "Prospecting" }, uid("OP")); }
  function updateOpportunity(id, vals) { opportunitiesCol.update(id, vals); }
  function deleteOpportunity(id) { opportunitiesCol.remove(id); }
  function addQuotation(vals) { quotationsCol.add({ client: vals.client, date: todayStr(), items: [{ id: vals.item, qty: vals.qty }], status: "Draft" }, uid("QT")); }
  function advanceQuotation(id) {
    const q = quotations.find((x) => x.id === id);
    const next = { Draft: "Sent", Sent: "Accepted" }[q?.status];
    if (next) quotationsCol.update(id, { status: next });
  }
  function convertQuotation(q) {
    const client = clientById[q.client];
    convertQuotationToOrder(q, client, prodById).catch(reportError);
  }
  function newSalesOrder(clientId, branchId, items) {
    salesOrdersCol.add({ client: clientId, branch: branchId, date: todayStr(), status: "Pending", items }, uid("SO"));
  }
  function advanceSalesOrder(id) {
    const order = salesOrders.find((o) => o.id === id);
    if (order.status === "Pending") {
      deliverSalesOrder(order, branchById[order.branch]?.name).catch(reportError);
      return;
    }
    const next = { Delivered: "Invoiced", Invoiced: "Paid" }[order.status];
    if (next) salesOrdersCol.update(id, { status: next });
  }
  function addReturn(vals) { returnsCol.add({ so: vals.so, item: vals.item, qty: vals.qty, reason: vals.reason, date: todayStr(), status: "Requested" }, uid("RT")); }
  function advanceReturn(r) {
    if (r.status === "Requested") { returnsCol.update(r.id, { status: "Approved" }); return; }
    if (r.status === "Approved") {
      const order = salesOrders.find((o) => o.id === r.so);
      restockReturn(r, order.branch).catch(reportError);
    }
  }

  /* ---- Actions: HR / Expenses / Projects / Service / Accounting ---- */
  function setLeaveStatus(id, status) { leavesCol.update(id, { status }); }
  function setExpenseStatus(id, status) { expensesCol.update(id, { status }); }
  function payPayroll(id) { payrollCol.update(id, { status: "Paid" }); }
  function addAttendance(vals) { attendanceCol.add(vals, uid("AT")); }
  function updateAttendance(id, vals) { attendanceCol.update(id, vals); }
  function deleteAttendance(id) { attendanceCol.remove(id); }
  function addLeave(vals) { leavesCol.add({ ...vals, status: "Pending" }, uid("LV")); }
  function addRecruit(vals) { recruitmentCol.add(vals, uid("REC")); }
  function updateRecruit(id, vals) { recruitmentCol.update(id, vals); }
  function deleteRecruit(id) { recruitmentCol.remove(id); }
  function addExpense(vals) { expensesCol.add({ ...vals, date: todayStr(), status: "Pending" }, uid("EXP")); }
  function updateExpense(id, vals) { expensesCol.update(id, vals); }
  function addProject(vals) { projectsCol.add({ ...vals, start: todayStr(), end: todayStr(), percent: 0 }, uid("PRJ")); }
  function updateProject(id, vals) { projectsCol.update(id, vals); }
  function addService(vals) { serviceCol.add({ ...vals, status: "Open", date: todayStr() }, uid("SV")); }
  function updateServiceTicket(id, vals) { serviceCol.update(id, vals); }
  function advanceService(id) {
    const s = service.find((x) => x.id === id);
    const next = { Open: "In Progress", "In Progress": "Resolved" }[s?.status];
    if (next) serviceCol.update(id, { status: next });
  }
  function addCoA(vals) { coaCol.add(vals, vals.code); }
  function updateCoA(code, vals) { coaCol.update(code, vals); }
  function deleteCoA(code) {
    const acct = coa.find((a) => a.code === code);
    if (journal.some((j) => j.debit === acct.name || j.credit === acct.name)) { alert("This account has journal entries posted against it — reassign or remove those first."); return; }
    coaCol.remove(code);
  }
  function addJournal(vals) { journalCol.add(vals, uid("JE")); }
  function updateJournal(id, vals) { journalCol.update(id, vals); }

  /* ---- Actions: Suppliers / Users / Banks / Fixed Assets / Documents / Manual Receipts & Payments ---- */
  function addSupplier(vals) { suppliersCol.add(vals, uid("SUP")); }
  function updateSupplier(id, vals) { suppliersCol.update(id, vals); }
  function deleteSupplier(id) {
    if (materials.some((m) => m.supplier === id)) { alert("This supplier is linked to existing materials — reassign those materials before deleting it."); return; }
    suppliersCol.remove(id);
  }
  function addUser(vals) { usersCol.add(vals, uid("u")); }
  function updateUserRole(id, newRole) {
    const target = users.find((u) => u.id === id);
    if (!target) return;
    const remainingAdmins = users.filter((u) => u.role === "Admin" && u.id !== id).length;
    if (target.role === "Admin" && newRole !== "Admin" && remainingAdmins === 0) {
      alert("You can't change the last Admin's role — promote someone else first.");
      return;
    }
    usersCol.update(id, { role: newRole });
  }
  function deleteUser(id) {
    const target = users.find((u) => u.id === id);
    if (!target) return;
    if (user && target.id === user.id) { alert("You can't delete the account you're currently signed in as."); return; }
    const remainingAdmins = users.filter((u) => u.role === "Admin" && u.id !== id).length;
    if (target.role === "Admin" && remainingAdmins === 0) { alert("You can't delete the last Admin account."); return; }
    if (!confirm(`Delete user "${target.name}"? This can't be undone.`)) return;
    usersCol.remove(id);
  }
  function addBank(vals) { banksCol.add(vals, uid("BANK")); }
  function updateBank(id, vals) { banksCol.update(id, vals); }
  function deleteBank(id) { banksCol.remove(id); }
  function addFixedAsset(vals) { fixedAssetsCol.add(vals, uid("FA")); }
  function updateFixedAsset(id, vals) { fixedAssetsCol.update(id, vals); }
  function deleteFixedAsset(id) { fixedAssetsCol.remove(id); }
  function addDocument(vals) { documentsCol.add({ ...vals, uploadedBy: user?.name || "Unknown", date: todayStr() }, uid("DOC")); }
  function updateDocument(id, vals) { documentsCol.update(id, vals); }
  function deleteDocument(id) { documentsCol.remove(id); }
  function addManualReceipt(vals) { manualReceiptsCol.add({ ...vals, date: todayStr() }, uid("RCP")); }
  function updateManualReceipt(id, vals) { manualReceiptsCol.update(id, vals); }
  function deleteManualReceipt(id) { manualReceiptsCol.remove(id); }
  function addManualPayment(vals) { manualPaymentsCol.add({ ...vals, date: todayStr() }, uid("PAY")); }
  function updateManualPayment(id, vals) { manualPaymentsCol.update(id, vals); }
  function deleteManualPayment(id) { manualPaymentsCol.remove(id); }

  /* ---- Actions: Taxes ---- */
  function addTaxRate(vals) { taxRatesCol.add({ ...vals, isActive: true }, uid("TAX")); }
  function updateTaxRate(id, vals) { taxRatesCol.update(id, vals); }
  function deleteTaxRate(id) { taxRatesCol.remove(id); }

  const ctxValue = {
    user, role, canEdit, isAdmin, money, currency, currencies, toUSD, t, lang, dir,
    materials, products, customers, suppliers, leads, opportunities, quotations, salesOrders, returns,
    purchaseRequests, purchaseOrders, purchaseInvoices, productionOrders, transfers, movements, stockCounts, itemLimits,
    coa, journal, banks, employees, attendance, leaves, payroll, recruitment, expenses, fixedAssets, projects, service, documents,
    users, manualReceipts, manualPayments,
    matById, prodById, clientById, supById, branchById, totalStock, orderTotal, poTotal, updateCustomer, updateEmployee,
    addPR, updatePR, approvePR, convertPRtoPO, newPO, receivePO, payPurchaseInvoice,
    canRunBatch, completeBatch, newBatch, transferStock, addStockCount,
    workCenters, addWorkCenter, updateWorkCenter, deleteWorkCenter,
    qualityChecks, addQualityCheck, updateQualityCheck, deleteQualityCheck,
    maintenance, addMaintenance, updateMaintenance, deleteMaintenance, logMaintenanceService,
    addStoreItem, removeStoreItem, upsertLimit, deleteItemLimit, addCurrency, updateCurrencyRate,
    addLead, updateLead, deleteLead, addOpportunity, updateOpportunity, deleteOpportunity,
    addQuotation, advanceQuotation, convertQuotation, newSalesOrder, advanceSalesOrder, addReturn, advanceReturn,
    setLeaveStatus, setExpenseStatus, updateExpense, payPayroll, addAttendance, updateAttendance, deleteAttendance,
    addLeave, addRecruit, updateRecruit, deleteRecruit, addExpense, addProject, updateProject, addService, updateServiceTicket, advanceService,
    addCoA, updateCoA, deleteCoA, addJournal, updateJournal,
    addSupplier, updateSupplier, deleteSupplier, addUser, updateUserRole, deleteUser, addBank, updateBank, deleteBank,
    addFixedAsset, updateFixedAsset, deleteFixedAsset, addDocument, updateDocument, deleteDocument,
    addManualReceipt, updateManualReceipt, deleteManualReceipt, addManualPayment, updateManualPayment, deleteManualPayment,
    addTaxRate, updateTaxRate, deleteTaxRate, taxRates,
  };

  if (stillLoading) {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: "#1B2421", color: "#F3EFE6" }}>
        <div className="text-center">
          <div className="text-sm uppercase tracking-widest opacity-70 mb-2">CodePulse</div>
          <div className="text-xs opacity-50">{firstError ? `Couldn't connect to Firebase: ${firstError.message}` : "Loading live data…"}</div>
        </div>
      </div>
    );
  }

  if (!user) return <Login onLogin={setUser} lang={lang} setLang={setLang} t={t} users={users} />;

  const allowedModules = MODULES.filter((m) => MODULE_ACCESS[m.id].includes(role));
  const activeModule = allowedModules.find((m) => m.id === selectedModule) ? selectedModule : allowedModules[0].id;
  const modDef = MODULES.find((m) => m.id === activeModule);
  const activeSub = modDef.subs ? (modDef.subs.find((s) => s.id === selectedSub) ? selectedSub : modDef.subs[0].id) : null;

  function selectModule(id) {
    setSelectedModule(id);
    const m = MODULES.find((x) => x.id === id);
    setSelectedSub(m.subs ? m.subs[0].id : null);
  }

  return (
    <Ctx.Provider value={ctxValue}>
      <div dir={dir} className="w-full min-h-[700px] flex" style={{ ...THEME_TOKENS[theme], background: "var(--bg-page)", fontFamily: "'Inter', system-ui, sans-serif" }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');`}</style>

        <div className="w-60 shrink-0 flex flex-col overflow-y-auto" style={{ background: "#1B2421" }}>
          <div className="px-5 pt-6 pb-5" style={{ borderBottom: "1px solid #2C362F" }}>
            <div className="flex items-center gap-2.5">
              <img src={LOGO_DATA_URI} alt="CodePulse" style={{ height: 30, width: "auto", borderRadius: 4 }} />
              <div className="text-lg font-bold tracking-wide" style={{ color: "#F3EFE6", fontFamily: "'Oswald', sans-serif" }}>CODEPULSE</div>
            </div>
            <div className="text-[10px] uppercase tracking-widest mt-1" style={{ color: "#C08A2E" }}>{t("Enterprise Console")}</div>
          </div>
          <nav className="flex-1 py-3">
            {allowedModules.map((m) => {
              const active = activeModule === m.id;
              return (
                <div key={m.id}>
                  <button onClick={() => selectModule(m.id)}
                    className="w-full text-left px-5 py-2 flex items-center gap-2 text-sm"
                    style={{ color: active ? "#F3EFE6" : "#9CA79E", background: active ? "#2C362F" : "transparent",
                      borderLeft: active ? "3px solid #C08A2E" : "3px solid transparent", fontWeight: active ? 600 : 500 }}>
                    {t(m.label)}
                  </button>
                  {active && m.subs && (
                    <div className="pb-1">
                      {m.subs.map((s) => (
                        <button key={s.id} onClick={() => setSelectedSub(s.id)}
                          className="w-full text-left pl-9 pr-5 py-1.5 text-xs"
                          style={{ color: activeSub === s.id ? "#F3EFE6" : "#7C8880", fontWeight: activeSub === s.id ? 600 : 400 }}>
                          {t(s.label)}
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </nav>
          <div className="px-5 py-4 text-[10px]" style={{ color: "#6F7C71", borderTop: "1px solid #2C362F" }}>
            {t("Line status:")} <span style={{ color: "#7EBF98" }}>● {t("Running")}</span>
          </div>
        </div>

        <div className="flex-1 flex flex-col">
          <div className="flex items-center justify-between px-6 py-3" style={{ borderBottom: "1px solid var(--border)", background: "var(--surface-alt)" }}>
            <div className="flex items-center gap-2">
              <span className="text-sm font-semibold" style={{ color: "var(--heading)" }}>{user.name}</span>
              <RolePill role={user.role} />
            </div>
            <div className="flex items-center gap-3">
              <div className="flex items-center gap-1.5">
                <span className="text-[11px] uppercase tracking-wide" style={{ color: "var(--text-label)" }}>{t("Language")}</span>
                <select value={lang} onChange={(e) => setLang(e.target.value)} className="text-xs px-2 py-1 rounded-sm" style={{ border: "1px solid var(--border-strong)" }}>
                  {LANGUAGES.map((l) => <option key={l.code} value={l.code}>{l.label}</option>)}
                </select>
              </div>
              <div className="flex items-center gap-1.5">
                <span className="text-[11px] uppercase tracking-wide" style={{ color: "var(--text-label)" }}>{t("Currency")}</span>
                <select value={currency} onChange={(e) => setCurrency(e.target.value)} className="text-xs px-2 py-1 rounded-sm font-mono" style={{ border: "1px solid var(--border-strong)" }}>
                  {currencies.map((c) => <option key={c.code} value={c.code}>{c.code}</option>)}
                </select>
              </div>
              <div className="flex items-center gap-1.5">
                <span className="text-[11px] uppercase tracking-wide" style={{ color: "var(--text-label)" }}>{t("Theme")}</span>
                <select value={theme} onChange={(e) => setTheme(e.target.value)} className="text-xs px-2 py-1 rounded-sm" style={{ border: "1px solid var(--border-strong)" }}>
                  <option value="light">{t("Light")}</option>
                  <option value="dark">{t("Dark")}</option>
                </select>
              </div>
              <Button variant="ghost" onClick={() => setUser(null)}>Sign Out</Button>
            </div>
          </div>
          <div className="flex-1 p-6 overflow-y-auto">
            <PageRouter module={activeModule} sub={activeSub} />
          </div>
        </div>
      </div>
    </Ctx.Provider>
  );
}

/* ================================================================
   Page Router
   ================================================================ */

function PageRouter({ module, sub }) {
  if (module === "dashboard") return <Dashboard />;
  if (module === "crm") {
    if (sub === "leads") return <LeadsPage />;
    if (sub === "customers") return <CustomersPage />;
    if (sub === "opportunities") return <OpportunitiesPage />;
  }
  if (module === "sales") {
    if (sub === "quotations") return <QuotationsPage />;
    if (sub === "salesorders") return <SalesOrdersPage />;
    if (sub === "delivery") return <DeliveryPage />;
    if (sub === "invoices") return <SalesInvoicesPage />;
    if (sub === "returns") return <ReturnsPage />;
  }
  if (module === "purchasing") {
    if (sub === "suppliers") return <SuppliersPage />;
    if (sub === "purchaserequests") return <PurchaseRequestsPage />;
    if (sub === "purchaseorders") return <PurchaseOrdersPage />;
    if (sub === "receipts") return <ReceiptsPage />;
    if (sub === "purchaseinvoices") return <PurchaseInvoicesPage />;
  }
  if (module === "inventory") {
    if (sub === "items") return <ItemsPage />;
    if (sub === "warehouses") return <WarehousesPage />;
    if (sub === "stockin") return <StockMovePage type="IN" />;
    if (sub === "stockout") return <StockMovePage type="OUT" />;
    if (sub === "transfers") return <TransfersPage />;
    if (sub === "stockcount") return <StockCountPage />;
  }
  if (module === "storemgmt") {
    if (sub === "additems") return <AddItemsPage />;
    if (sub === "outitems") return <OutItemsPage />;
    if (sub === "sm_transfers") return <TransfersPage />;
    if (sub === "itemlimits") return <ItemLimitsPage />;
  }
  if (module === "accounting") {
    if (sub === "coa") return <CoaPage />;
    if (sub === "journal") return <JournalPage />;
    if (sub === "cash") return <CashPage />;
    if (sub === "banks") return <BanksPage />;
    if (sub === "arreceipts") return <ArReceiptsPage />;
    if (sub === "appayments") return <ApPaymentsPage />;
    if (sub === "taxes") return <TaxesPage />;
    if (sub === "finreports") return <FinReportsPage />;
  }
  if (module === "hr") {
    if (sub === "employees") return <EmployeesPage />;
    if (sub === "attendance") return <AttendancePage />;
    if (sub === "leaves") return <LeavesPage />;
    if (sub === "payroll") return <PayrollPage />;
    if (sub === "recruitment") return <RecruitmentPage />;
  }
  if (module === "expenses") return <ExpensesPage />;
  if (module === "fixedassets") return <FixedAssetsPage />;
  if (module === "manufacturing") {
    if (sub === "productionorders") return <ManufacturingPage />;
    if (sub === "workcenters") return <WorkCentersPage />;
    if (sub === "qualitycontrol") return <QualityControlPage />;
    if (sub === "maintenance") return <MaintenancePage />;
  }
  if (module === "projects") return <ProjectsPage />;
  if (module === "service") return <ServicePage />;
  if (module === "documents") return <DocumentsPage />;
  if (module === "reports") return <ReportsPage />;
  if (module === "approvals") return <ApprovalsPage />;
  if (module === "users") return <UsersPage />;
  if (module === "settings") {
    if (sub === "currencies") return <CurrenciesPage />;
    return <SettingsPage />;
  }
  return null;
}

/* ================================================================
   Dashboard
   ================================================================ */

function Dashboard() {
  const { salesOrders, clientById, orderTotal, money, products, materials, totalStock, productionOrders } = useApp();
  const revenuePaid = salesOrders.filter((o) => o.status === "Paid").reduce((s, o) => s + orderTotal(o), 0);
  const outstanding = salesOrders.filter((o) => o.status === "Invoiced").reduce((s, o) => s + orderTotal(o), 0);
  const openOrders = salesOrders.filter((o) => o.status === "Pending" || o.status === "Delivered").length;
  const activeBatches = productionOrders.filter((p) => p.status === "Planned").length;
  const lowStockMats = materials.filter((m) => m.stock < m.reorder);
  const lowStockProds = products.filter((p) => totalStock(p) < p.reorder);

  const byClient = {};
  salesOrders.forEach((o) => { byClient[o.client] = (byClient[o.client] || 0) + orderTotal(o); });
  const maxClientRev = Math.max(1, ...Object.values(byClient));

  return (
    <div>
      <SectionHeader title="Dashboard" sub="Enterprise snapshot across CRM, sales, production, and finance." />
      <div className="grid grid-cols-4 gap-4 mb-6">
        <Card title="Revenue Collected" value={money(revenuePaid)} sub="Paid invoices" accent="#3F7D5C" />
        <Card title="Outstanding" value={money(outstanding)} sub="Invoiced, unpaid" accent="#8A6A2E" />
        <Card title="Open Sales Orders" value={openOrders} sub="Pending / delivered" />
        <Card title="Active Batches" value={activeBatches} sub="On the production line" accent="#3A5C86" />
      </div>
      <div className="grid grid-cols-2 gap-4">
        <div className="p-4 rounded-md" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
          <div className="text-sm font-semibold mb-3" style={{ color: "var(--heading)" }}>Low stock alerts</div>
          {lowStockMats.length === 0 && lowStockProds.length === 0 && <div className="text-sm" style={{ color: "var(--text-secondary)" }}>All stock levels are healthy.</div>}
          <div className="space-y-3">
            {lowStockMats.map((m) => <Gauge key={m.id} label={`${m.name} (raw material)`} value={m.stock} max={m.reorder * 1.5} />)}
            {lowStockProds.map((p) => <Gauge key={p.id} label={`${p.name} (finished good)`} value={totalStock(p)} max={p.reorder * 1.5} />)}
          </div>
        </div>
        <div className="p-4 rounded-md" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
          <div className="text-sm font-semibold mb-3" style={{ color: "var(--heading)" }}>Revenue by hotel account</div>
          <div className="space-y-2.5">
            {Object.entries(byClient).map(([cid, rev]) => (
              <div key={cid}>
                <div className="flex justify-between text-xs mb-1" style={{ color: "var(--text)" }}><span>{clientById[cid]?.name}</span><span className="font-mono">{money(rev)}</span></div>
                <div className="h-2 rounded-full" style={{ background: "var(--border)" }}><div className="h-2 rounded-full" style={{ width: `${(rev / maxClientRev) * 100}%`, background: "#C08A2E" }} /></div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ================================================================
   CRM
   ================================================================ */

function LeadsPage() {
  const { leads, addLead, updateLead, deleteLead, canEdit } = useApp();
  return (
    <ListPage title="Leads" sub="Prospective hotel accounts not yet converted to customers." canEdit={canEdit}
      addFields={[{ key: "name", label: "Hotel Name" }, { key: "company", label: "Group" }, { key: "source", label: "Source" }, { key: "owner", label: "Owner", default: "Jordan Blake" }]}
      onAdd={addLead} onUpdate={updateLead} onDelete={canEdit ? deleteLead : undefined}
      columns={["ID", "Hotel", "Group", "Source", "Owner", "Status", "Date"]} rows={leads}
      renderRow={(l) => (<><Td mono>{l.id}</Td><Td>{l.name}</Td><Td>{l.company}</Td><Td>{l.source}</Td><Td>{l.owner}</Td><Td><Pill>{l.status}</Pill></Td><Td mono>{l.date}</Td></>)} />
  );
}
function CustomersPage() {
  const { customers, salesOrders, orderTotal, branchById, money, updateCustomer, isAdmin } = useApp();
  const [editingId, setEditingId] = useState(null);
  const editingRow = customers.find((c) => c.id === editingId);
  const fields = [{ key: "name", label: "Hotel Name" }, { key: "location", label: "Location" }, { key: "contact", label: "Contact" }, { key: "branch", label: "Servicing Branch", type: "select", options: Object.values(branchById).map((b) => ({ value: b.id, label: b.name })) }];
  return (
    <div>
      <SectionHeader title="Customers" sub="Hotel accounts with active or historical orders." />
      {editingRow && (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updateCustomer(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      )}
      <Table title="Customers" columns={["ID", "Hotel", "Location", "Servicing Branch", "Contact", "Lifetime Orders", "Lifetime Value", ""]} rows={customers}
        renderRow={(c) => {
          const orders = salesOrders.filter((o) => o.client === c.id);
          const value = orders.reduce((s, o) => s + orderTotal(o), 0);
          return (<><Td mono>{c.id}</Td><Td>{c.name}</Td><Td>{c.location}</Td><Td>{branchById[c.branch].name}</Td><Td>{c.contact}</Td><Td mono>{orders.length}</Td><Td mono>{money(value)}</Td>
            <Td>{isAdmin && <Button variant="ghost" onClick={() => setEditingId(c.id)}>Edit</Button>}</Td></>);
        }} />
    </div>
  );
}
function OpportunitiesPage() {
  const { opportunities, addOpportunity, updateOpportunity, deleteOpportunity, canEdit, clientById, money, customers } = useApp();
  return (
    <ListPage title="Opportunities" sub="Deals in progress with existing or prospective hotel accounts." canEdit={canEdit}
      addFields={[{ key: "customer", label: "Customer", type: "select", options: customers.map((c) => ({ value: c.id, label: c.name })) }, { key: "title", label: "Deal" }, { key: "value", label: "Value", type: "money", default: 5000 }, { key: "closeDate", label: "Close Date", type: "date", default: "2026-09-30" }]}
      onAdd={addOpportunity} onUpdate={updateOpportunity} onDelete={canEdit ? deleteOpportunity : undefined}
      columns={["ID", "Customer", "Deal", "Value", "Stage", "Close Date"]} rows={opportunities}
      renderRow={(o) => (<><Td mono>{o.id}</Td><Td>{clientById[o.customer]?.name || o.customer}</Td><Td>{o.title}</Td><Td mono>{money(o.value)}</Td><Td><Pill>{o.stage}</Pill></Td><Td mono>{o.closeDate}</Td></>)} />
  );
}

/* ================================================================
   Sales
   ================================================================ */

function QuotationsPage() {
  const { quotations, customers, prodById, products, addQuotation, advanceQuotation, convertQuotation, canEdit, clientById } = useApp();
  const NEXT = { Draft: "Send", Sent: "Mark Accepted" };
  return (
    <div>
      <SectionHeader title="Quotations" sub="Price quotes sent to hotels before a firm order is placed." />
      {canEdit && <AddForm
        fields={[{ key: "client", label: "Client", type: "select", options: customers.map((c) => ({ value: c.id, label: c.name })) }, { key: "item", label: "Product", type: "select", options: products.map((p) => ({ value: p.id, label: p.name })) }, { key: "qty", label: "Qty", type: "number", default: 100 }]}
        onSubmit={addQuotation} submitLabel="+ Create Quotation" />}
      <Table title="Quotations" columns={["Quote", "Client", "Date", "Items", "Status", ""]} rows={quotations}
        renderRow={(q) => (
          <>
            <Td mono>{q.id}</Td><Td>{clientById[q.client].name}</Td><Td mono>{q.date}</Td>
            <Td><div className="text-xs space-y-0.5">{q.items.map((it, i) => <div key={i}>{it.qty} × {prodById[it.id].name}</div>)}</div></Td>
            <Td><Pill>{q.status}</Pill></Td>
            <Td>
              {canEdit && NEXT[q.status] && <Button onClick={() => advanceQuotation(q.id)}>{NEXT[q.status]}</Button>}
              {canEdit && q.status === "Accepted" && <Button variant="accent" onClick={() => convertQuotation(q)}>Convert to Order</Button>}
            </Td>
          </>
        )} />
    </div>
  );
}
function SalesOrdersPage() {
  const { salesOrders, clientById, prodById, orderTotal, money } = useApp();
  return (
    <div>
      <SectionHeader title="Sales Orders" sub="Confirmed orders from hotel clients, from intake through payment." />
      <Table title="Sales Orders" columns={["Order", "Client", "Date", "Items", "Total", "Status"]} rows={salesOrders}
        renderRow={(o) => (
          <>
            <Td mono>{o.id}</Td><Td>{clientById[o.client].name}</Td><Td mono>{o.date}</Td>
            <Td><div className="text-xs space-y-0.5">{o.items.map((it, i) => <div key={i}>{it.qty} × {prodById[it.id].name}</div>)}</div></Td>
            <Td mono>{money(orderTotal(o))}</Td><Td><Pill>{o.status}</Pill></Td>
          </>
        )} />
    </div>
  );
}
function DeliveryPage() {
  const { salesOrders, clientById, branchById, prodById, advanceSalesOrder, canEdit } = useApp();
  return (
    <div>
      <SectionHeader title="Delivery" sub="Ship confirmed orders from their servicing branch." />
      <Table title="Delivery" columns={["Order", "Client", "Branch", "Items", "Status", ""]} rows={salesOrders}
        renderRow={(o) => (
          <>
            <Td mono>{o.id}</Td><Td>{clientById[o.client].name}</Td><Td>{branchById[o.branch].name}</Td>
            <Td><div className="text-xs space-y-0.5">{o.items.map((it, i) => <div key={i}>{it.qty} × {prodById[it.id].name}</div>)}</div></Td>
            <Td><Pill>{o.status}</Pill></Td>
            <Td>{canEdit && o.status === "Pending" && <Button onClick={() => advanceSalesOrder(o.id)}>Mark Delivered</Button>}</Td>
          </>
        )} />
    </div>
  );
}
function SalesInvoicesPage() {
  const { salesOrders, clientById, orderTotal, advanceSalesOrder, money, canEdit } = useApp();
  const relevant = salesOrders.filter((o) => o.status === "Delivered" || o.status === "Invoiced" || o.status === "Paid");
  const NEXT = { Delivered: "Create Invoice", Invoiced: "Mark Paid" };
  return (
    <div>
      <SectionHeader title="Invoices" sub="Billing status for every delivered hotel order." />
      <Table title="Invoices" columns={["Invoice", "Client", "Date", "Amount", "Status", ""]} rows={relevant}
        renderRow={(o) => (
          <>
            <Td mono>{o.id.replace("SO", "INV")}</Td><Td>{clientById[o.client].name}</Td><Td mono>{o.date}</Td>
            <Td mono>{money(orderTotal(o))}</Td><Td><Pill>{o.status}</Pill></Td>
            <Td>{canEdit && NEXT[o.status] && <Button onClick={() => advanceSalesOrder(o.id)}>{NEXT[o.status]}</Button>}</Td>
          </>
        )} />
      {relevant.length === 0 && <div className="text-sm mt-4" style={{ color: "var(--text-secondary)" }}>No invoices yet — deliver a sales order first.</div>}
    </div>
  );
}
function ReturnsPage() {
  const { returns, salesOrders, prodById, addReturn, advanceReturn, canEdit } = useApp();
  const eligible = salesOrders.filter((o) => ["Delivered", "Invoiced", "Paid"].includes(o.status));
  const NEXT = { Requested: "Approve", Approved: "Restock" };
  return (
    <div>
      <SectionHeader title="Returns" sub="Return requests against delivered orders." />
      {canEdit && <AddForm
        fields={[{ key: "so", label: "Sales Order", type: "select", options: eligible.map((o) => ({ value: o.id, label: o.id })) }, { key: "item", label: "Item", type: "select", options: Object.values(prodById).map((p) => ({ value: p.id, label: p.name })) }, { key: "qty", label: "Qty", type: "number", default: 5 }, { key: "reason", label: "Reason", default: "Damaged in transit" }]}
        onSubmit={addReturn} submitLabel="+ Request Return" />}
      <Table title="Returns" columns={["Return", "Order", "Item", "Qty", "Reason", "Date", "Status", ""]} rows={returns}
        renderRow={(r) => (
          <>
            <Td mono>{r.id}</Td><Td mono>{r.so}</Td><Td>{prodById[r.item]?.name}</Td><Td mono>{r.qty}</Td><Td>{r.reason}</Td><Td mono>{r.date}</Td>
            <Td><Pill>{r.status}</Pill></Td>
            <Td>{canEdit && NEXT[r.status] && <Button onClick={() => advanceReturn(r)}>{NEXT[r.status]}</Button>}</Td>
          </>
        )} />
    </div>
  );
}

/* ================================================================
   Purchasing
   ================================================================ */

function SuppliersPage() {
  const { suppliers, addSupplier, updateSupplier, deleteSupplier, canEdit, isAdmin } = useApp();
  const [editingId, setEditingId] = useState(null);
  const fields = [{ key: "name", label: "Name" }, { key: "contact", label: "Contact" }, { key: "lead", label: "Lead Time", type: "number", default: 5 }];
  const editingRow = suppliers.find((s) => s.id === editingId);
  return (
    <div>
      <SectionHeader title="Suppliers" sub="Vendors providing raw materials to the factory." />
      {editingRow ? (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updateSupplier(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      ) : canEdit && (
        <AddForm fields={fields} onSubmit={addSupplier} submitLabel="+ Add Supplier" />
      )}
      <Table title="Suppliers" columns={["ID", "Supplier", "Contact", "Lead Time", ""]} rows={suppliers}
        renderRow={(s) => (<><Td mono>{s.id}</Td><Td>{s.name}</Td><Td>{s.contact}</Td><Td mono>{s.lead} days</Td>
          <Td>{isAdmin && <Button variant="ghost" onClick={() => setEditingId(s.id)}>Edit</Button>}{canEdit && <Button variant="danger" onClick={() => deleteSupplier(s.id)}>Delete</Button>}</Td></>)} />
    </div>
  );
}
function PurchaseRequestsPage() {
  const { purchaseRequests, materials, matById, addPR, updatePR, approvePR, convertPRtoPO, canEdit, isAdmin } = useApp();
  const [editingId, setEditingId] = useState(null);
  const fields = [{ key: "item", label: "Material", type: "select", options: materials.map((m) => ({ value: m.id, label: m.name })) }, { key: "qty", label: "Qty", type: "number", default: 100 }, { key: "requestedBy", label: "Requested By", default: "Priya Nair" }, { key: "department", label: "Department", default: "Production" }];
  const editingRow = purchaseRequests.find((p) => p.id === editingId);
  return (
    <div>
      <SectionHeader title="Purchase Requests" sub="Internal requests for raw materials, approved before a PO is raised." />
      {editingRow ? (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updatePR(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      ) : canEdit && <AddForm fields={fields} onSubmit={addPR} submitLabel="+ Request" />}
      <Table title="Purchase Requests" columns={["PR", "Material", "Qty", "Requested By", "Department", "Date", "Status", ""]} rows={purchaseRequests}
        renderRow={(p) => (
          <>
            <Td mono>{p.id}</Td><Td>{matById[p.item].name}</Td><Td mono>{p.qty} {matById[p.item].unit}</Td><Td>{p.requestedBy}</Td><Td>{p.department}</Td><Td mono>{p.date}</Td>
            <Td><Pill>{p.status}</Pill></Td>
            <Td>
              {canEdit && p.status === "Pending" && <><Button onClick={() => approvePR(p.id, "Approved")}>Approve</Button><Button variant="danger" onClick={() => approvePR(p.id, "Rejected")}>Reject</Button></>}
              {canEdit && p.status === "Approved" && <Button variant="accent" onClick={() => convertPRtoPO(p)}>Convert to PO</Button>}
              {isAdmin && <Button variant="ghost" onClick={() => setEditingId(p.id)}>Edit</Button>}
            </Td>
          </>
        )} />
    </div>
  );
}
function PurchaseOrdersPage() {
  const { purchaseOrders, materials, suppliers, supById, matById, poTotal, receivePO, newPO, money, canEdit } = useApp();
  const [supplierId, setSupplierId] = useState(suppliers[0].id);
  const [materialId, setMaterialId] = useState(materials[0].id);
  const [qty, setQty] = useState(200);
  return (
    <div>
      <SectionHeader title="Purchase Orders" sub="Confirmed orders sent to suppliers for raw materials." />
      {canEdit && (
        <div className="p-4 rounded-md mb-5 flex items-end gap-3 flex-wrap" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
          <div><div className="text-[11px] uppercase tracking-wide mb-1" style={{ color: "var(--text-label)" }}>Supplier</div>
            <select value={supplierId} onChange={(e) => setSupplierId(e.target.value)} className="text-sm px-2 py-1.5 rounded-sm" style={{ border: "1px solid var(--border-strong)" }}>
              {suppliers.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select></div>
          <div><div className="text-[11px] uppercase tracking-wide mb-1" style={{ color: "var(--text-label)" }}>Material</div>
            <select value={materialId} onChange={(e) => setMaterialId(e.target.value)} className="text-sm px-2 py-1.5 rounded-sm" style={{ border: "1px solid var(--border-strong)" }}>
              {materials.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}
            </select></div>
          <div><div className="text-[11px] uppercase tracking-wide mb-1" style={{ color: "var(--text-label)" }}>Qty</div>
            <input type="number" value={qty} onChange={(e) => setQty(+e.target.value)} className="text-sm px-2 py-1.5 rounded-sm w-24" style={{ border: "1px solid var(--border-strong)" }} /></div>
          <Button variant="accent" onClick={() => newPO(supplierId, [{ id: materialId, qty, cost: matById[materialId].cost }])}>+ Create PO</Button>
        </div>
      )}
      <Table title="Purchase Orders" columns={["PO", "Supplier", "Date", "Items", "Total", "Status", ""]} rows={purchaseOrders}
        renderRow={(po) => (
          <>
            <Td mono>{po.id}</Td><Td>{supById[po.supplier].name}</Td><Td mono>{po.date}</Td>
            <Td><div className="text-xs space-y-0.5">{po.items.map((it, i) => <div key={i}>{it.qty} {matById[it.id].unit} · {matById[it.id].name}</div>)}</div></Td>
            <Td mono>{money(poTotal(po))}</Td><Td><Pill>{po.status}</Pill></Td>
            <Td>{canEdit && po.status === "Pending" && <Button onClick={() => receivePO(po.id)}>Receive</Button>}</Td>
          </>
        )} />
    </div>
  );
}
function ReceiptsPage() {
  const { purchaseOrders, supById, matById } = useApp();
  const received = purchaseOrders.filter((p) => p.status === "Received");
  return (
    <div>
      <SectionHeader title="Receipts" sub="Goods receipt notes for materials received into the factory." />
      <Table title="Receipts" columns={["Receipt", "PO", "Supplier", "Date", "Items"]} rows={received}
        renderRow={(po) => (
          <>
            <Td mono>{po.id.replace("PO", "GRN")}</Td><Td mono>{po.id}</Td><Td>{supById[po.supplier].name}</Td><Td mono>{po.date}</Td>
            <Td><div className="text-xs space-y-0.5">{po.items.map((it, i) => <div key={i}>{it.qty} {matById[it.id].unit} · {matById[it.id].name}</div>)}</div></Td>
          </>
        )} />
      {received.length === 0 && <div className="text-sm mt-4" style={{ color: "var(--text-secondary)" }}>No goods received yet.</div>}
    </div>
  );
}
function PurchaseInvoicesPage() {
  const { purchaseInvoices, supById, money, payPurchaseInvoice, canEdit } = useApp();
  return (
    <div>
      <SectionHeader title="Invoices" sub="Supplier invoices generated when goods are received." />
      <Table title="Invoices" columns={["Invoice", "PO", "Supplier", "Amount", "Date", "Status", ""]} rows={purchaseInvoices}
        renderRow={(pi) => (
          <>
            <Td mono>{pi.id}</Td><Td mono>{pi.po}</Td><Td>{supById[pi.supplier].name}</Td><Td mono>{money(pi.amount)}</Td><Td mono>{pi.date}</Td>
            <Td><Pill>{pi.status}</Pill></Td>
            <Td>{canEdit && pi.status === "Unpaid" && <Button onClick={() => payPurchaseInvoice(pi.id)}>Mark Paid</Button>}</Td>
          </>
        )} />
    </div>
  );
}

/* ================================================================
   Inventory
   ================================================================ */

function ItemsPage() {
  const { materials, products, totalStock, money } = useApp();
  const items = [
    ...materials.map((m) => ({ id: m.id, name: m.name, type: "Raw Material", unit: m.unit, stock: m.stock, reorder: m.reorder, value: m.cost })),
    ...products.map((p) => ({ id: p.id, name: p.name, type: "Finished Good", unit: p.unit, stock: totalStock(p), reorder: p.reorder, value: p.price })),
  ];
  return (
    <div>
      <SectionHeader title="Items" sub="Unified catalog of raw materials and finished goods." />
      <Table title="Items" columns={["ID", "Item", "Type", "Stock", "Reorder At", "Unit Value"]} rows={items}
        renderRow={(it) => (
          <>
            <Td mono>{it.id}</Td><Td>{it.name}</Td><Td><Pill>{it.type}</Pill></Td>
            <Td mono>{it.stock} {it.unit}{it.stock < it.reorder && <span className="ml-2"><Pill>Pending</Pill></span>}</Td>
            <Td mono>{it.reorder} {it.unit}</Td><Td mono>{money(it.value)}</Td>
          </>
        )} />
    </div>
  );
}
function WarehousesPage() {
  const { products } = useApp();
  return (
    <div>
      <SectionHeader title="Warehouses" sub="Factory and store locations holding inventory." />
      <div className="grid grid-cols-3 gap-4 mb-6">
        {BRANCHES.map((b) => (
          <div key={b.id} className="p-4 rounded-md" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
            <div className="flex items-center justify-between mb-1"><div className="text-sm font-semibold" style={{ color: "var(--heading)" }}>{b.name}</div><Pill>{b.type}</Pill></div>
            <div className="text-xs" style={{ color: "var(--text-secondary)" }}>{b.location}</div>
          </div>
        ))}
      </div>
      <div className="text-sm font-semibold mb-2" style={{ color: "var(--heading)" }}>Finished goods stock by warehouse</div>
      <Table title="Warehouses" columns={["Item", "HQ Factory", "Denver Store", "Tampa Store"]} rows={products}
        renderRow={(p) => (<><Td>{p.name}</Td><Td mono>{p.stockByBranch["BR-HQ"]}</Td><Td mono>{p.stockByBranch["BR-DEN"]}</Td><Td mono>{p.stockByBranch["BR-TPA"]}</Td></>)} />
    </div>
  );
}
function StockMovePage({ type }) {
  const { movements, matById, prodById, branchById } = useApp();
  const rows = movements.filter((m) => m.type === type).slice().reverse();
  return (
    <div>
      <SectionHeader title={type === "IN" ? "Stock In" : "Stock Out"} sub={type === "IN" ? "Inbound movements from receiving and production." : "Outbound movements from production consumption and deliveries."} />
      <Table title={type === "IN" ? "Stock In" : "Stock Out"} columns={["Movement", "Item", "Type", "Qty", "Warehouse", "Date", "Reference"]} rows={rows}
        renderRow={(m) => {
          const name = m.itemType === "Material" ? matById[m.item]?.name : prodById[m.item]?.name;
          const unit = m.itemType === "Material" ? matById[m.item]?.unit : prodById[m.item]?.unit;
          return (<><Td mono>{m.id}</Td><Td>{name}</Td><Td>{m.itemType}</Td><Td mono>{m.qty} {unit}</Td><Td>{branchById[m.warehouse].name}</Td><Td mono>{m.date}</Td><Td mono>{m.ref}</Td></>);
        }} />
    </div>
  );
}
function TransfersPage() {
  const { products, transfers, branchById, transferStock, canEdit } = useApp();
  const [productId, setProductId] = useState(products[0].id);
  const [from, setFrom] = useState("BR-HQ");
  const [to, setTo] = useState(BRANCHES.find((b) => b.id !== "BR-HQ").id);
  const [qty, setQty] = useState(50);
  return (
    <div>
      <SectionHeader title="Transfers" sub="Move finished goods between warehouses." />
      {canEdit && (
        <div className="p-4 rounded-md mb-5 flex items-end gap-3 flex-wrap" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
          <div><div className="text-[11px] uppercase tracking-wide mb-1" style={{ color: "var(--text-label)" }}>Product</div>
            <select value={productId} onChange={(e) => setProductId(e.target.value)} className="text-sm px-2 py-1.5 rounded-sm" style={{ border: "1px solid var(--border-strong)" }}>
              {products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
            </select></div>
          <div><div className="text-[11px] uppercase tracking-wide mb-1" style={{ color: "var(--text-label)" }}>From</div>
            <select value={from} onChange={(e) => setFrom(e.target.value)} className="text-sm px-2 py-1.5 rounded-sm" style={{ border: "1px solid var(--border-strong)" }}>
              {BRANCHES.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}
            </select></div>
          <div><div className="text-[11px] uppercase tracking-wide mb-1" style={{ color: "var(--text-label)" }}>To</div>
            <select value={to} onChange={(e) => setTo(e.target.value)} className="text-sm px-2 py-1.5 rounded-sm" style={{ border: "1px solid var(--border-strong)" }}>
              {BRANCHES.filter((b) => b.id !== from).map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}
            </select></div>
          <div><div className="text-[11px] uppercase tracking-wide mb-1" style={{ color: "var(--text-label)" }}>Qty</div>
            <input type="number" value={qty} onChange={(e) => setQty(+e.target.value)} className="text-sm px-2 py-1.5 rounded-sm w-24" style={{ border: "1px solid var(--border-strong)" }} /></div>
          <Button variant="accent" onClick={() => transferStock(productId, from, to, qty)}>Transfer Stock</Button>
        </div>
      )}
      <Table title="Transfers" columns={["Transfer", "Product", "From", "To", "Qty", "Date"]} rows={transfers}
        renderRow={(t) => {
          const p = products.find((x) => x.id === t.product);
          return (<><Td mono>{t.id}</Td><Td>{p?.name}</Td><Td>{branchById[t.from].name}</Td><Td>{branchById[t.to].name}</Td><Td mono>{t.qty}</Td><Td mono>{t.date}</Td></>);
        }} />
    </div>
  );
}
function StockCountPage() {
  const { stockCounts, materials, products, matById, prodById, branchById, addStockCount, canEdit } = useApp();
  const itemOptions = [...materials.map((m) => ({ value: m.id, label: m.name })), ...products.map((p) => ({ value: p.id, label: p.name }))];
  return (
    <div>
      <SectionHeader title="Stock Count" sub="Physical inventory counts reconciled against system quantities." />
      {canEdit && <AddForm
        fields={[{ key: "warehouse", label: "Warehouse", type: "select", options: BRANCHES.map((b) => ({ value: b.id, label: b.name })) }, { key: "itemId", label: "Item", type: "select", options: itemOptions }, { key: "countedQty", label: "Counted Qty", type: "number", default: 0 }]}
        onSubmit={addStockCount} submitLabel="+ Record Count" />}
      <Table title="Stock Count" columns={["Count", "Warehouse", "Item", "System Qty", "Counted Qty", "Variance", "Date", "Status"]} rows={stockCounts}
        renderRow={(sc) => {
          const name = sc.item.startsWith("RM") ? matById[sc.item]?.name : prodById[sc.item]?.name;
          const variance = sc.countedQty - sc.systemQty;
          return (
            <>
              <Td mono>{sc.id}</Td><Td>{branchById[sc.warehouse].name}</Td><Td>{name}</Td><Td mono>{sc.systemQty}</Td><Td mono>{sc.countedQty}</Td>
              <Td mono><span style={{ color: variance === 0 ? undefined : variance < 0 ? "#A64B3A" : "#3F7D5C" }}>{variance > 0 ? "+" : ""}{variance}</span></Td>
              <Td mono>{sc.date}</Td><Td><Pill>{sc.status}</Pill></Td>
            </>
          );
        }} />
    </div>
  );
}

/* ================================================================
   Store Management
   ================================================================ */

function AddItemsPage() {
  const { materials, products, addStoreItem, canEdit } = useApp();
  const itemOptions = [...materials.map((m) => ({ value: m.id, label: `${m.name} (Raw Material)` })), ...products.map((p) => ({ value: p.id, label: `${p.name} (Finished Good)` }))];
  return (
    <div>
      <SectionHeader title="Add Items" sub="Bring stock into a warehouse directly — new intake, found stock, or corrections." />
      {canEdit && <AddForm
        fields={[{ key: "itemId", label: "Item", type: "select", options: itemOptions }, { key: "warehouse", label: "Warehouse", type: "select", options: BRANCHES.map((b) => ({ value: b.id, label: b.name })) }, { key: "qty", label: "Qty", type: "number", default: 50 }, { key: "reason", label: "Reason", default: "Manual Add" }]}
        onSubmit={addStoreItem} submitLabel="+ Add Stock" />}
      <div className="text-sm" style={{ color: "var(--text-secondary)" }}>Every addition here is logged in Inventory → Stock In with the reason you give it.</div>
    </div>
  );
}
function OutItemsPage() {
  const { materials, products, removeStoreItem, canEdit } = useApp();
  const itemOptions = [...materials.map((m) => ({ value: m.id, label: `${m.name} (Raw Material)` })), ...products.map((p) => ({ value: p.id, label: `${p.name} (Finished Good)` }))];
  return (
    <div>
      <SectionHeader title="Out Items" sub="Remove stock from a warehouse for reasons outside a normal sale — damage, loss, samples, internal use." />
      {canEdit && <AddForm
        fields={[{ key: "itemId", label: "Item", type: "select", options: itemOptions }, { key: "warehouse", label: "Warehouse", type: "select", options: BRANCHES.map((b) => ({ value: b.id, label: b.name })) }, { key: "qty", label: "Qty", type: "number", default: 10 }, { key: "reason", label: "Reason", type: "select", options: ["Damaged", "Lost", "Sample", "Internal Use", "Other"] }]}
        onSubmit={removeStoreItem} submitLabel="− Remove Stock" />}
      <div className="text-sm" style={{ color: "var(--text-secondary)" }}>Every removal here is logged in Inventory → Stock Out with the reason you give it.</div>
    </div>
  );
}
function ItemLimitsPage() {
  const { itemLimits, materials, products, matById, prodById, branchById, upsertLimit, deleteItemLimit, canEdit, isAdmin } = useApp();
  const itemOptions = [...materials.map((m) => ({ value: m.id, label: m.name })), ...products.map((p) => ({ value: p.id, label: p.name }))];
  const [editingId, setEditingId] = useState(null);
  function currentStock(itemId, warehouse) {
    if (itemId.startsWith("RM")) return warehouse === "BR-HQ" ? matById[itemId].stock : 0;
    return prodById[itemId].stockByBranch[warehouse];
  }
  const fields = [{ key: "item", label: "Item", type: "select", options: itemOptions }, { key: "warehouse", label: "Warehouse", type: "select", options: BRANCHES.map((b) => ({ value: b.id, label: b.name })) }, { key: "min", label: "Min", type: "number", default: 50 }, { key: "max", label: "Max", type: "number", default: 500 }];
  const editingRow = itemLimits.find((l) => l.id === editingId);
  return (
    <div>
      <SectionHeader title="Item Limits" sub="Set a minimum and maximum stock level per item, per warehouse." />
      {editingRow ? (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { upsertLimit(vals); setEditingId(null); }} submitLabel="Save Changes" />
      ) : canEdit && <AddForm fields={fields} onSubmit={upsertLimit} submitLabel="Set Limit" />}
      <Table title="Item Limits" columns={["ID", "Item", "Warehouse", "Current Stock", "Min", "Max", "Status", ""]} rows={itemLimits}
        renderRow={(l) => {
          const name = l.item.startsWith("RM") ? matById[l.item]?.name : prodById[l.item]?.name;
          const stock = currentStock(l.item, l.warehouse);
          const status = stock < l.min ? "Below Min" : stock > l.max ? "Above Max" : "Within Range";
          return (<><Td mono>{l.id}</Td><Td>{name}</Td><Td>{branchById[l.warehouse].name}</Td><Td mono>{stock}</Td><Td mono>{l.min}</Td><Td mono>{l.max}</Td><Td><Pill>{status}</Pill></Td>
            <Td>{isAdmin && <Button variant="ghost" onClick={() => setEditingId(l.id)}>Edit</Button>}{canEdit && <Button variant="danger" onClick={() => deleteItemLimit(l.id)}>Delete</Button>}</Td></>);
        }} />
      {itemLimits.length === 0 && <div className="text-sm mt-3" style={{ color: "var(--text-secondary)" }}>No limits set yet.</div>}
    </div>
  );
}

/* ================================================================
   Accounting
   ================================================================ */

function CoaPage() {
  const { coa, addCoA, updateCoA, deleteCoA, canEdit } = useApp();
  return (
    <ListPage title="Chart of Accounts" sub="The ledger accounts used across journal entries and reports." canEdit={canEdit} idKey="code"
      addFields={[{ key: "code", label: "Code", default: "6000" }, { key: "name", label: "Account Name" }, { key: "type", label: "Type", type: "select", options: ["Asset", "Liability", "Equity", "Revenue", "Expense"] }]}
      onAdd={addCoA} onUpdate={updateCoA} onDelete={canEdit ? deleteCoA : undefined}
      columns={["Code", "Account", "Type"]} rows={coa}
      renderRow={(a) => (<><Td mono>{a.code}</Td><Td>{a.name}</Td><Td><Pill>{a.type}</Pill></Td></>)} />
  );
}
function JournalPage() {
  const { journal, coa, addJournal, updateJournal, canEdit, isAdmin, money } = useApp();
  const [editingId, setEditingId] = useState(null);
  const fields = [{ key: "date", label: "Date", type: "date", default: "2026-08-17" }, { key: "memo", label: "Memo" }, { key: "debit", label: "Debit Account", type: "select", options: coa.map((a) => a.name) }, { key: "credit", label: "Credit Account", type: "select", options: coa.map((a) => a.name) }, { key: "amount", label: "Amount", type: "money", default: 100 }];
  const editingRow = journal.find((j) => j.id === editingId);
  return (
    <div>
      <SectionHeader title="Journal Entries" sub="Manual double-entry postings." />
      {editingRow ? (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updateJournal(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      ) : canEdit && <AddForm fields={fields} onSubmit={addJournal} />}
      <Table title="Journal Entries" columns={["Entry", "Date", "Memo", "Debit", "Credit", "Amount", ""]} rows={journal}
        renderRow={(j) => (<><Td mono>{j.id}</Td><Td mono>{j.date}</Td><Td>{j.memo}</Td><Td>{j.debit}</Td><Td>{j.credit}</Td><Td mono>{money(j.amount)}</Td>
          <Td>{isAdmin && <Button variant="ghost" onClick={() => setEditingId(j.id)}>Edit</Button>}</Td></>)} />
    </div>
  );
}
function CashPage() {
  const { salesOrders, purchaseInvoices, expenses, clientById, supById, orderTotal, money } = useApp();
  const entries = [
    ...salesOrders.filter((o) => o.status === "Paid").map((o) => ({ date: o.date, desc: `Receipt — ${clientById[o.client].name} (${o.id})`, inflow: orderTotal(o), outflow: 0 })),
    ...purchaseInvoices.filter((p) => p.status === "Paid").map((p) => ({ date: p.date, desc: `Payment — ${supById[p.supplier].name} (${p.po})`, inflow: 0, outflow: p.amount })),
    ...expenses.filter((e) => e.status === "Paid").map((e) => ({ date: e.date, desc: `Expense — ${e.category} (${e.employee})`, inflow: 0, outflow: e.amount })),
  ].sort((a, b) => a.date.localeCompare(b.date));
  let bal = 15000;
  const withBalance = entries.map((e) => { bal = bal + e.inflow - e.outflow; return { ...e, balance: bal }; });
  return (
    <div>
      <SectionHeader title="Cash" sub="Simplified cash ledger — opening balance $15,000." />
      <Table title="Cash" columns={["Date", "Description", "Inflow", "Outflow", "Balance"]} rows={withBalance}
        renderRow={(e) => (<><Td mono>{e.date}</Td><Td>{e.desc}</Td><Td mono>{e.inflow ? money(e.inflow) : "—"}</Td><Td mono>{e.outflow ? money(e.outflow) : "—"}</Td><Td mono>{money(e.balance)}</Td></>)} />
    </div>
  );
}
function BanksPage() {
  const { banks, addBank, updateBank, deleteBank, money, canEdit, isAdmin } = useApp();
  const [editingId, setEditingId] = useState(null);
  const fields = [{ key: "name", label: "Bank Account" }, { key: "number", label: "Number", default: "****0000" }, { key: "balance", label: "Balance", type: "money", default: 0 }];
  const editingRow = banks.find((b) => b.id === editingId);
  return (
    <div>
      <SectionHeader title="Banks" sub="Company bank accounts." />
      {editingRow ? (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updateBank(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      ) : canEdit && (
        <AddForm fields={fields} onSubmit={addBank} submitLabel="+ Add Bank Account" />
      )}
      <Table title="Banks" columns={["ID", "Bank Account", "Number", "Balance", ""]} rows={banks}
        renderRow={(b) => (<><Td mono>{b.id}</Td><Td>{b.name}</Td><Td mono>{b.number}</Td><Td mono>{money(b.balance)}</Td>
          <Td>{isAdmin && <Button variant="ghost" onClick={() => setEditingId(b.id)}>Edit</Button>}{canEdit && <Button variant="danger" onClick={() => deleteBank(b.id)}>Delete</Button>}</Td></>)} />
    </div>
  );
}
function ArReceiptsPage() {
  const { salesOrders, clientById, orderTotal, money, customers, manualReceipts, addManualReceipt, updateManualReceipt, deleteManualReceipt, canEdit, isAdmin } = useApp();
  const [editingId, setEditingId] = useState(null);
  const fromOrders = salesOrders.filter((o) => o.status === "Paid").map((o) => ({
    id: o.id.replace("SO", "RCP"), party: clientById[o.client].name, date: o.date, amount: orderTotal(o), source: "Sales Order",
  }));
  const manual = manualReceipts.map((r) => ({ id: r.id, party: clientById[r.client]?.name || r.client, date: r.date, amount: r.amount, source: "Manual" }));
  const rows = [...fromOrders, ...manual];
  const fields = [{ key: "client", label: "Client", type: "select", options: customers.map((c) => ({ value: c.id, label: c.name })) }, { key: "amount", label: "Amount", type: "money", default: 100 }];
  const editingRow = manualReceipts.find((r) => r.id === editingId);
  return (
    <div>
      <SectionHeader title="Receipts" sub="Customer payments received — automatic from paid sales orders, or recorded manually." />
      {editingRow ? (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updateManualReceipt(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      ) : canEdit && (
        <AddForm fields={fields} onSubmit={addManualReceipt} submitLabel="+ Record Receipt" />
      )}
      <Table title="Receipts" columns={["Receipt", "Client", "Date", "Amount", "Source", ""]} rows={rows}
        renderRow={(r) => (<><Td mono>{r.id}</Td><Td>{r.party}</Td><Td mono>{r.date}</Td><Td mono>{money(r.amount)}</Td><Td><Pill>{r.source}</Pill></Td>
          <Td>{isAdmin && r.source === "Manual" && <Button variant="ghost" onClick={() => setEditingId(r.id)}>Edit</Button>}{canEdit && r.source === "Manual" && <Button variant="danger" onClick={() => deleteManualReceipt(r.id)}>Delete</Button>}</Td></>)} />
      {rows.length === 0 && <div className="text-sm mt-4" style={{ color: "var(--text-secondary)" }}>No receipts yet.</div>}
    </div>
  );
}
function ApPaymentsPage() {
  const { purchaseInvoices, expenses, supById, money, suppliers, manualPayments, addManualPayment, updateManualPayment, deleteManualPayment, canEdit, isAdmin } = useApp();
  const [editingId, setEditingId] = useState(null);
  const rows = [
    ...purchaseInvoices.filter((p) => p.status === "Paid").map((p) => ({ id: p.id, payee: supById[p.supplier].name, date: p.date, amount: p.amount, type: "Supplier" })),
    ...expenses.filter((e) => e.status === "Paid").map((e) => ({ id: e.id, payee: e.employee, date: e.date, amount: e.amount, type: "Expense" })),
    ...manualPayments.map((p) => ({ id: p.id, payee: p.payee, date: p.date, amount: p.amount, type: "Manual" })),
  ];
  const fields = [{ key: "payee", label: "Payee", type: "select", options: suppliers.map((s) => ({ value: s.name, label: s.name })) }, { key: "amount", label: "Amount", type: "money", default: 100 }];
  const editingRow = manualPayments.find((p) => p.id === editingId);
  return (
    <div>
      <SectionHeader title="Payments" sub="Payments made to suppliers and reimbursed expenses — or recorded manually." />
      {editingRow ? (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updateManualPayment(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      ) : canEdit && (
        <AddForm fields={fields} onSubmit={addManualPayment} submitLabel="+ Record Payment" />
      )}
      <Table title="Payments" columns={["Ref", "Payee", "Type", "Date", "Amount", ""]} rows={rows}
        renderRow={(r) => (<><Td mono>{r.id}</Td><Td>{r.payee}</Td><Td><Pill>{r.type}</Pill></Td><Td mono>{r.date}</Td><Td mono>{money(r.amount)}</Td>
          <Td>{isAdmin && r.type === "Manual" && <Button variant="ghost" onClick={() => setEditingId(r.id)}>Edit</Button>}{canEdit && r.type === "Manual" && <Button variant="danger" onClick={() => deleteManualPayment(r.id)}>Delete</Button>}</Td></>)} />
    </div>
  );
}
function TaxesPage() {
  const { taxRates, addTaxRate, updateTaxRate, deleteTaxRate, canEdit, isAdmin } = useApp();
  const [editingId, setEditingId] = useState(null);
  const fields = [
    { key: "name", label: "Tax Name", default: "New Tax Rate" },
    { key: "rate", label: "Rate %", type: "number", default: 5 },
    { key: "type", label: "Type", type: "select", options: ["Sales Tax", "VAT", "Duty", "Withholding", "Other"] },
    { key: "appliesTo", label: "Applies To", type: "select", options: ["Sales", "Purchases", "Both"] },
  ];
  const editingRow = taxRates.find((r) => r.id === editingId);
  const activeCount = taxRates.filter((r) => r.isActive).length;
  const avgRate = taxRates.length ? (taxRates.reduce((s, r) => s + r.rate, 0) / taxRates.length).toFixed(2) : "0.00";

  return (
    <div>
      <SectionHeader title="Taxes" sub="Tax rates applied across sales and purchasing — configure once, reference everywhere." />

      <div className="grid grid-cols-3 gap-4 mb-6">
        <Card title="Tax Rates Configured" value={String(taxRates.length)} />
        <Card title="Active Rates" value={String(activeCount)} accent="#3F7D5C" />
        <Card title="Average Rate" value={`${avgRate}%`} accent="#C08A2E" />
      </div>

      {editingRow ? (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updateTaxRate(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      ) : canEdit && (
        <AddForm fields={fields} onSubmit={addTaxRate} submitLabel="+ Add Tax Rate" />
      )}

      <Table title="Taxes" columns={["ID", "Tax Name", "Rate", "Type", "Applies To", "Status", ""]} rows={taxRates}
        renderRow={(r) => (
          <>
            <Td mono>{r.id}</Td><Td>{r.name}</Td><Td mono>{r.rate.toFixed(2)}%</Td><Td>{r.type}</Td><Td>{r.appliesTo}</Td>
            <Td><Pill>{r.isActive ? "Active" : "Inactive"}</Pill></Td>
            <Td>
              {canEdit && <Button variant="ghost" onClick={() => updateTaxRate(r.id, { isActive: !r.isActive })}>{r.isActive ? "Deactivate" : "Activate"}</Button>}
              {isAdmin && <Button variant="ghost" onClick={() => setEditingId(r.id)}>Edit</Button>}
              {isAdmin && <Button variant="danger" onClick={() => deleteTaxRate(r.id)}>Delete</Button>}
            </Td>
          </>
        )} />
      <div className="text-xs mt-3" style={{ color: "var(--text-label)" }}>
        Tax rates are configured here for reference and reporting; applying them automatically to individual invoice line items is a natural next step once that's needed.
      </div>
    </div>
  );
}

function FinReportsPage() {
  const { salesOrders, purchaseInvoices, expenses, payroll, orderTotal, money } = useApp();
  const revenue = salesOrders.filter((o) => ["Invoiced", "Paid"].includes(o.status)).reduce((s, o) => s + orderTotal(o), 0);
  const cogs = purchaseInvoices.reduce((s, p) => s + p.amount, 0);
  const opex = expenses.filter((e) => ["Approved", "Paid"].includes(e.status)).reduce((s, e) => s + e.amount, 0);
  const payrollCost = payroll.reduce((s, p) => s + (p.base - p.deductions), 0);
  const net = revenue - cogs - opex - payrollCost;
  return (
    <div>
      <SectionHeader title="Financial Reports" sub="Simplified profit & loss summary — not a full accrual statement." />
      <div className="grid grid-cols-2 gap-4 max-w-xl">
        <Card title="Revenue" value={money(revenue)} sub="Invoiced + paid sales orders" accent="#3F7D5C" />
        <Card title="Cost of Goods Sold" value={money(cogs)} sub="Supplier invoices" accent="#A64B3A" />
        <Card title="Operating Expenses" value={money(opex)} sub="Approved + paid expense claims" accent="#A64B3A" />
        <Card title="Payroll Cost" value={money(payrollCost)} sub="Net pay across current run" accent="#A64B3A" />
      </div>
      <div className="mt-5 p-4 rounded-md max-w-xl" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
        <div className="flex justify-between text-sm font-semibold" style={{ color: "var(--heading)" }}><span>Net Profit</span><span className="font-mono" style={{ color: net >= 0 ? "#3F7D5C" : "#A64B3A" }}>{money(net)}</span></div>
      </div>
    </div>
  );
}

/* ================================================================
   HR
   ================================================================ */

function EmployeesPage() {
  const { employees, branchById, updateEmployee, isAdmin } = useApp();
  const [editingId, setEditingId] = useState(null);
  const editingRow = employees.find((e) => e.id === editingId);
  const fields = [{ key: "name", label: "Name" }, { key: "role", label: "Role" }, { key: "department", label: "Department" }, { key: "branch", label: "Branch", type: "select", options: Object.values(branchById).map((b) => ({ value: b.id, label: b.name })) }, { key: "email", label: "Email" }, { key: "status", label: "Status", type: "select", options: ["Active", "On Leave", "Inactive"] }];
  return (
    <div>
      <SectionHeader title="Employees" sub="Factory and store staff." />
      {editingRow && (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updateEmployee(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      )}
      <Table title="Employees" columns={["ID", "Name", "Role", "Department", "Branch", "Email", "Hired", "Status", ""]} rows={employees}
        renderRow={(e) => (<><Td mono>{e.id}</Td><Td>{e.name}</Td><Td>{e.role}</Td><Td>{e.department}</Td><Td>{branchById[e.branch].name}</Td><Td>{e.email}</Td><Td mono>{e.hired}</Td><Td><Pill>{e.status}</Pill></Td>
          <Td>{isAdmin && <Button variant="ghost" onClick={() => setEditingId(e.id)}>Edit</Button>}</Td></>)} />
    </div>
  );
}
function AttendancePage() {
  const { attendance, employees, addAttendance, updateAttendance, deleteAttendance, canEdit } = useApp();
  return (
    <ListPage title="Attendance" sub="Daily attendance log." canEdit={canEdit}
      addFields={[{ key: "employee", label: "Employee", type: "select", options: employees.map((e) => e.name) }, { key: "date", label: "Date", type: "date", default: "2026-08-17" }, { key: "status", label: "Status", type: "select", options: ["Present", "Absent", "Late"] }]}
      onAdd={addAttendance} onUpdate={updateAttendance} onDelete={canEdit ? deleteAttendance : undefined}
      columns={["ID", "Employee", "Date", "Status"]} rows={attendance}
      renderRow={(a) => (<><Td mono>{a.id}</Td><Td>{a.employee}</Td><Td mono>{a.date}</Td><Td><Pill>{a.status}</Pill></Td></>)} />
  );
}
function LeavesPage() {
  const { leaves, employees, addLeave, setLeaveStatus, canEdit } = useApp();
  return (
    <div>
      <SectionHeader title="Leaves" sub="Employee leave requests." />
      {canEdit && <AddForm fields={[{ key: "employee", label: "Employee", type: "select", options: employees.map((e) => e.name) }, { key: "type", label: "Type", type: "select", options: ["Sick", "Vacation", "Personal", "Unpaid"] }, { key: "from", label: "From", type: "date", default: "2026-08-20" }, { key: "to", label: "To", type: "date", default: "2026-08-21" }]} onSubmit={addLeave} submitLabel="+ Request Leave" />}
      <Table title="Leaves" columns={["ID", "Employee", "Type", "From", "To", "Status", ""]} rows={leaves}
        renderRow={(l) => (
          <>
            <Td mono>{l.id}</Td><Td>{l.employee}</Td><Td>{l.type}</Td><Td mono>{l.from}</Td><Td mono>{l.to}</Td><Td><Pill>{l.status}</Pill></Td>
            <Td>{canEdit && l.status === "Pending" && <><Button onClick={() => setLeaveStatus(l.id, "Approved")}>Approve</Button><Button variant="danger" onClick={() => setLeaveStatus(l.id, "Rejected")}>Reject</Button></>}</Td>
          </>
        )} />
    </div>
  );
}
function PayrollPage() {
  const { payroll, payPayroll, money, canEdit } = useApp();
  return (
    <div>
      <SectionHeader title="Payroll" sub="Monthly payroll run." />
      <Table title="Payroll" columns={["Run", "Employee", "Month", "Base", "Deductions", "Net Pay", "Status", ""]} rows={payroll}
        renderRow={(p) => {
          const net = p.base - p.deductions;
          return (
            <>
              <Td mono>{p.id}</Td><Td>{p.employee}</Td><Td mono>{p.month}</Td><Td mono>{money(p.base)}</Td><Td mono>{money(p.deductions)}</Td><Td mono>{money(net)}</Td>
              <Td><Pill>{p.status}</Pill></Td>
              <Td>{canEdit && p.status === "Draft" && <Button onClick={() => payPayroll(p.id)}>Mark Paid</Button>}</Td>
            </>
          );
        }} />
    </div>
  );
}
function RecruitmentPage() {
  const { recruitment, addRecruit, updateRecruit, deleteRecruit, canEdit } = useApp();
  return (
    <ListPage title="Recruitment" sub="Open positions and candidate pipeline." canEdit={canEdit}
      addFields={[{ key: "position", label: "Position" }, { key: "department", label: "Department" }, { key: "candidate", label: "Candidate" }, { key: "stage", label: "Stage", type: "select", options: ["Applied", "Interview", "Offer", "Hired", "Rejected"] }]}
      onAdd={addRecruit} onUpdate={updateRecruit} onDelete={canEdit ? deleteRecruit : undefined}
      columns={["ID", "Position", "Department", "Candidate", "Stage"]} rows={recruitment}
      renderRow={(r) => (<><Td mono>{r.id}</Td><Td>{r.position}</Td><Td>{r.department}</Td><Td>{r.candidate}</Td><Td><Pill>{r.stage}</Pill></Td></>)} />
  );
}

/* ================================================================
   Expenses / Fixed Assets / Manufacturing / Projects / Service / Documents
   ================================================================ */

function ExpensesPage() {
  const { expenses, employees, addExpense, updateExpense, setExpenseStatus, money, canEdit, isAdmin } = useApp();
  const [editingId, setEditingId] = useState(null);
  const fields = [{ key: "employee", label: "Employee", type: "select", options: employees.map((e) => e.name) }, { key: "category", label: "Category" }, { key: "amount", label: "Amount", type: "money", default: 100 }];
  const editingRow = expenses.find((e) => e.id === editingId);
  return (
    <div>
      <SectionHeader title="Expenses" sub="Employee expense claims." />
      {editingRow ? (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updateExpense(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      ) : canEdit && <AddForm fields={fields} onSubmit={addExpense} submitLabel="+ Submit Claim" />}
      <Table title="Expenses" columns={["ID", "Employee", "Category", "Amount", "Date", "Status", ""]} rows={expenses}
        renderRow={(e) => (
          <>
            <Td mono>{e.id}</Td><Td>{e.employee}</Td><Td>{e.category}</Td><Td mono>{money(e.amount)}</Td><Td mono>{e.date}</Td><Td><Pill>{e.status}</Pill></Td>
            <Td>
              {canEdit && e.status === "Pending" && <><Button onClick={() => setExpenseStatus(e.id, "Approved")}>Approve</Button><Button variant="danger" onClick={() => setExpenseStatus(e.id, "Rejected")}>Reject</Button></>}
              {canEdit && e.status === "Approved" && <Button onClick={() => setExpenseStatus(e.id, "Paid")}>Mark Paid</Button>}
              {isAdmin && <Button variant="ghost" onClick={() => setEditingId(e.id)}>Edit</Button>}
            </Td>
          </>
        )} />
    </div>
  );
}
function FixedAssetsPage() {
  const { fixedAssets, addFixedAsset, updateFixedAsset, deleteFixedAsset, money, canEdit, isAdmin } = useApp();
  const today = new Date("2026-08-17");
  const [editingId, setEditingId] = useState(null);
  const fields = [{ key: "name", label: "Asset" }, { key: "category", label: "Category" }, { key: "purchaseDate", label: "Purchase Date", type: "date", default: "2026-08-17" }, { key: "cost", label: "Cost", type: "money", default: 5000 }, { key: "lifeYears", label: "Useful Life (yrs)", type: "number", default: 5 }];
  const editingRow = fixedAssets.find((a) => a.id === editingId);
  return (
    <div>
      <SectionHeader title="Fixed Assets" sub="Asset register with straight-line depreciation." />
      {editingRow ? (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updateFixedAsset(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      ) : canEdit && (
        <AddForm fields={fields} onSubmit={addFixedAsset} submitLabel="+ Add Asset" />
      )}
      <Table title="Fixed Assets" columns={["ID", "Asset", "Category", "Purchase Date", "Cost", "Useful Life", "Book Value", ""]} rows={fixedAssets}
        renderRow={(a) => {
          const yrs = (today - new Date(a.purchaseDate)) / (365.25 * 24 * 3600 * 1000);
          const annualDep = a.cost / a.lifeYears;
          const accumulated = Math.min(a.cost, annualDep * Math.max(0, yrs));
          const book = Math.max(0, a.cost - accumulated);
          return (<><Td mono>{a.id}</Td><Td>{a.name}</Td><Td>{a.category}</Td><Td mono>{a.purchaseDate}</Td><Td mono>{money(a.cost)}</Td><Td mono>{a.lifeYears} yrs</Td><Td mono>{money(book)}</Td>
            <Td>{isAdmin && <Button variant="ghost" onClick={() => setEditingId(a.id)}>Edit</Button>}{canEdit && <Button variant="danger" onClick={() => deleteFixedAsset(a.id)}>Delete</Button>}</Td></>);
        }} />
    </div>
  );
}
function ManufacturingPage() {
  const { productionOrders, products, materials, matById, canRunBatch, completeBatch, newBatch, canEdit } = useApp();
  const [productId, setProductId] = useState(products[0].id);
  const [qty, setQty] = useState(100);

  const activeBatches = productionOrders.filter((b) => b.status === "Planned");
  const completedBatches = productionOrders.filter((b) => b.status === "Completed");
  const atRiskBatches = activeBatches.filter((b) => !canRunBatch(products.find((p) => p.id === b.product), b.qty));
  const unitsInProduction = activeBatches.reduce((s, b) => s + b.qty, 0);

  return (
    <div>
      <SectionHeader title="Production Orders" sub="Plan batches at the factory, track material consumption, and complete runs." />

      <div className="grid grid-cols-4 gap-4 mb-6">
        <Card title="Active Batches" value={String(activeBatches.length)} accent="#3A5C86" />
        <Card title="Units In Production" value={unitsInProduction.toLocaleString()} />
        <Card title="Completed Batches" value={String(completedBatches.length)} accent="#3F7D5C" />
        <Card title="Batches At Risk" value={String(atRiskBatches.length)} sub={atRiskBatches.length ? "Short on materials" : "All clear"} accent={atRiskBatches.length ? "#A64B3A" : "#3F7D5C"} />
      </div>

      {canEdit && (
        <div className="p-4 rounded-md mb-5 flex items-end gap-3" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
          <div><div className="text-[11px] uppercase tracking-wide mb-1" style={{ color: "var(--text-label)" }}>Product</div>
            <select value={productId} onChange={(e) => setProductId(e.target.value)} className="text-sm px-2 py-1.5 rounded-sm" style={{ border: "1px solid var(--border-strong)" }}>
              {products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
            </select></div>
          <div><div className="text-[11px] uppercase tracking-wide mb-1" style={{ color: "var(--text-label)" }}>Batch Qty</div>
            <input type="number" value={qty} onChange={(e) => setQty(+e.target.value)} className="text-sm px-2 py-1.5 rounded-sm w-24" style={{ border: "1px solid var(--border-strong)" }} /></div>
          <Button variant="accent" onClick={() => newBatch(productId, qty)}>+ Schedule Batch</Button>
        </div>
      )}
      <Table title="Manufacturing" columns={["Batch", "Product", "Qty", "Date", "Status", "Materials Needed", ""]} rows={productionOrders}
        renderRow={(b) => {
          const product = products.find((p) => p.id === b.product);
          const ok = canRunBatch(product, b.qty);
          return (
            <>
              <Td mono>{b.id}</Td><Td>{product.name}</Td><Td mono>{b.qty} {product.unit}</Td><Td mono>{b.date}</Td>
              <Td><Pill>{b.status}</Pill></Td>
              <Td><div className="text-xs space-y-0.5">{product.bom.map((line) => (
                <div key={line.id} style={{ color: matById[line.id].stock >= line.qty * b.qty ? "var(--text-secondary)" : "#A64B3A" }}>{matById[line.id].name}: {(line.qty * b.qty).toFixed(1)} {matById[line.id].unit}</div>
              ))}</div></Td>
              <Td>{canEdit && b.status === "Planned" && <Button onClick={() => completeBatch(b.id)} disabled={!ok}>{ok ? "Complete" : "Short stock"}</Button>}</Td>
            </>
          );
        }} />
    </div>
  );
}

function WorkCentersPage() {
  const { workCenters, addWorkCenter, updateWorkCenter, deleteWorkCenter, canEdit, isAdmin } = useApp();
  const [editingId, setEditingId] = useState(null);
  const fields = [
    { key: "name", label: "Work Center" },
    { key: "department", label: "Department", type: "select", options: ["Devices", "Linens", "Amenities"] },
    { key: "capacityPerDay", label: "Capacity / Day", type: "number", default: 100 },
    { key: "status", label: "Status", type: "select", options: ["Running", "Idle", "Maintenance"] },
  ];
  const editingRow = workCenters.find((w) => w.id === editingId);
  return (
    <div>
      <SectionHeader title="Work Centers" sub="Production lines and stations on the factory floor, and what they're rated to handle." />
      {editingRow ? (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updateWorkCenter(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      ) : canEdit && <AddForm fields={fields} onSubmit={addWorkCenter} submitLabel="+ Add Work Center" />}
      <Table title="Work Centers" columns={["ID", "Work Center", "Department", "Capacity / Day", "Status", ""]} rows={workCenters}
        renderRow={(w) => (
          <>
            <Td mono>{w.id}</Td><Td>{w.name}</Td><Td>{w.department}</Td><Td mono>{w.capacityPerDay}/day</Td><Td><Pill>{w.status}</Pill></Td>
            <Td>{isAdmin && <Button variant="ghost" onClick={() => setEditingId(w.id)}>Edit</Button>}{canEdit && <Button variant="danger" onClick={() => deleteWorkCenter(w.id)}>Delete</Button>}</Td>
          </>
        )} />
    </div>
  );
}

function QualityControlPage() {
  const { qualityChecks, productionOrders, products, addQualityCheck, updateQualityCheck, deleteQualityCheck, canEdit, isAdmin } = useApp();
  const [editingId, setEditingId] = useState(null);
  const fields = [
    { key: "batch", label: "Batch", type: "select", options: productionOrders.map((b) => ({ value: b.id, label: `${b.id} — ${products.find((p) => p.id === b.product)?.name}` })) },
    { key: "product", label: "Product", type: "select", options: products.map((p) => ({ value: p.id, label: p.name })) },
    { key: "inspector", label: "Inspector", default: "Aiko Tanaka" },
    { key: "result", label: "Result", type: "select", options: ["Pending", "Pass", "Fail"] },
    { key: "notes", label: "Notes", default: "" },
  ];
  const editingRow = qualityChecks.find((q) => q.id === editingId);
  const passRate = qualityChecks.length ? Math.round((qualityChecks.filter((q) => q.result === "Pass").length / qualityChecks.filter((q) => q.result !== "Pending").length || 0) * 100) : 0;
  return (
    <div>
      <SectionHeader title="Quality Control" sub="Inspection records for finished-goods batches, before they ship to hotel clients." />
      <div className="grid grid-cols-3 gap-4 mb-6">
        <Card title="Inspections Logged" value={String(qualityChecks.length)} />
        <Card title="Pass Rate" value={`${passRate}%`} accent={passRate >= 90 ? "#3F7D5C" : "#A64B3A"} />
        <Card title="Pending Review" value={String(qualityChecks.filter((q) => q.result === "Pending").length)} accent="#8A6A2E" />
      </div>
      {editingRow ? (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updateQualityCheck(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      ) : canEdit && <AddForm fields={fields} onSubmit={addQualityCheck} submitLabel="+ Log Inspection" />}
      <Table title="Quality Control" columns={["ID", "Batch", "Product", "Inspector", "Date", "Result", "Notes", ""]} rows={qualityChecks}
        renderRow={(q) => (
          <>
            <Td mono>{q.id}</Td><Td mono>{q.batch}</Td><Td>{products.find((p) => p.id === q.product)?.name}</Td><Td>{q.inspector}</Td><Td mono>{q.date}</Td>
            <Td><Pill>{q.result}</Pill></Td><Td className="text-xs">{q.notes}</Td>
            <Td>{isAdmin && <Button variant="ghost" onClick={() => setEditingId(q.id)}>Edit</Button>}{canEdit && <Button variant="danger" onClick={() => deleteQualityCheck(q.id)}>Delete</Button>}</Td>
          </>
        )} />
    </div>
  );
}

function MaintenancePage() {
  const { maintenance, workCenters, addMaintenance, updateMaintenance, deleteMaintenance, logMaintenanceService, canEdit, isAdmin } = useApp();
  const [editingId, setEditingId] = useState(null);
  const fields = [
    { key: "equipment", label: "Equipment" },
    { key: "workCenter", label: "Work Center", type: "select", options: workCenters.map((w) => ({ value: w.id, label: w.name })) },
    { key: "lastService", label: "Last Service", type: "date", default: "2026-08-17" },
    { key: "nextDue", label: "Next Due", type: "date", default: "2026-11-17" },
    { key: "technician", label: "Technician", default: "Diego Alvarez" },
  ];
  const editingRow = maintenance.find((m) => m.id === editingId);
  const overdueCount = maintenance.filter((m) => m.status === "Overdue").length;
  return (
    <div>
      <SectionHeader title="Maintenance" sub="Preventive maintenance schedule for factory equipment." />
      <div className="grid grid-cols-3 gap-4 mb-6">
        <Card title="Equipment Tracked" value={String(maintenance.length)} />
        <Card title="Due Soon" value={String(maintenance.filter((m) => m.status === "Due Soon").length)} accent="#8A6A2E" />
        <Card title="Overdue" value={String(overdueCount)} accent={overdueCount ? "#A64B3A" : "#3F7D5C"} sub={overdueCount ? "Needs attention" : "All clear"} />
      </div>
      {editingRow ? (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updateMaintenance(editingId, { ...vals, status: "Up to Date" }); setEditingId(null); }} submitLabel="Save Changes" />
      ) : canEdit && <AddForm fields={fields} onSubmit={(vals) => addMaintenance({ ...vals, status: "Up to Date" })} submitLabel="+ Add Equipment" />}
      <Table title="Maintenance" columns={["ID", "Equipment", "Work Center", "Last Service", "Next Due", "Technician", "Status", ""]} rows={maintenance}
        renderRow={(m) => (
          <>
            <Td mono>{m.id}</Td><Td>{m.equipment}</Td><Td>{workCenters.find((w) => w.id === m.workCenter)?.name}</Td>
            <Td mono>{m.lastService}</Td><Td mono>{m.nextDue}</Td><Td>{m.technician}</Td><Td><Pill>{m.status}</Pill></Td>
            <Td>
              {canEdit && m.status !== "Up to Date" && <Button variant="accent" onClick={() => logMaintenanceService(m.id)}>Log Service</Button>}
              {isAdmin && <Button variant="ghost" onClick={() => setEditingId(m.id)}>Edit</Button>}
              {canEdit && <Button variant="danger" onClick={() => deleteMaintenance(m.id)}>Delete</Button>}
            </Td>
          </>
        )} />
    </div>
  );
}

function ProjectsPage() {
  const { projects, addProject, updateProject, money, canEdit, isAdmin } = useApp();
  const [editingId, setEditingId] = useState(null);
  const fields = [{ key: "name", label: "Project" }, { key: "client", label: "Client", default: "Internal" }, { key: "manager", label: "Manager" }, { key: "status", label: "Status", type: "select", options: ["Planning", "In Progress", "On Hold", "Completed"] }, { key: "budget", label: "Budget", type: "money", default: 10000 }];
  const editingRow = projects.find((p) => p.id === editingId);
  return (
    <div>
      <SectionHeader title="Projects" sub="Internal and client-facing initiatives." />
      {editingRow ? (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updateProject(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      ) : canEdit && <AddForm fields={fields} onSubmit={addProject} submitLabel="+ New Project" />}
      <Table title="Projects" columns={["ID", "Project", "Client", "Manager", "Status", "Budget", "Progress", ""]} rows={projects}
        renderRow={(p) => (
          <>
            <Td mono>{p.id}</Td><Td>{p.name}</Td><Td>{p.client}</Td><Td>{p.manager}</Td><Td><Pill>{p.status}</Pill></Td><Td mono>{money(p.budget)}</Td>
            <Td><div className="w-24 h-2 rounded-full" style={{ background: "var(--border)" }}><div className="h-2 rounded-full" style={{ width: `${p.percent}%`, background: "#C08A2E" }} /></div></Td>
            <Td>{isAdmin && <Button variant="ghost" onClick={() => setEditingId(p.id)}>Edit</Button>}</Td>
          </>
        )} />
    </div>
  );
}
function ServicePage() {
  const { service, clientById, addService, updateServiceTicket, advanceService, customers, canEdit, isAdmin } = useApp();
  const NEXT = { Open: "Start Work", "In Progress": "Mark Resolved" };
  const [editingId, setEditingId] = useState(null);
  const fields = [{ key: "client", label: "Client", type: "select", options: customers.map((c) => ({ value: c.id, label: c.name })) }, { key: "issue", label: "Issue" }, { key: "priority", label: "Priority", type: "select", options: ["Low", "Medium", "High"] }, { key: "assignedTo", label: "Assigned To", default: "Priya Nair" }];
  const editingRow = service.find((s) => s.id === editingId);
  return (
    <div>
      <SectionHeader title="Service" sub="Post-sale support tickets from hotel clients." />
      {editingRow ? (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updateServiceTicket(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      ) : canEdit && <AddForm fields={fields} onSubmit={addService} submitLabel="+ New Ticket" />}
      <Table title="Service" columns={["Ticket", "Client", "Issue", "Priority", "Assigned To", "Date", "Status", ""]} rows={service}
        renderRow={(s) => (
          <>
            <Td mono>{s.id}</Td><Td>{clientById[s.client]?.name || s.client}</Td><Td>{s.issue}</Td><Td><Pill>{s.priority}</Pill></Td><Td>{s.assignedTo}</Td><Td mono>{s.date}</Td>
            <Td><Pill>{s.status}</Pill></Td>
            <Td>{canEdit && NEXT[s.status] && <Button onClick={() => advanceService(s.id)}>{NEXT[s.status]}</Button>}{isAdmin && <Button variant="ghost" onClick={() => setEditingId(s.id)}>Edit</Button>}</Td>
          </>
        )} />
    </div>
  );
}
function DocumentsPage() {
  const { documents, addDocument, updateDocument, deleteDocument, canEdit, isAdmin } = useApp();
  const [editingId, setEditingId] = useState(null);
  const fields = [{ key: "name", label: "File Name", default: "New Contract.pdf" }, { key: "module", label: "Module", type: "select", options: ["Sales", "Purchasing", "HR", "Accounting", "Legal", "Other"] }, { key: "type", label: "Type", type: "select", options: ["Contract", "Report", "Compliance", "Other"] }];
  const editingRow = documents.find((d) => d.id === editingId);
  return (
    <div>
      <SectionHeader title="Documents" sub="Contracts, reports, and compliance records." />
      {editingRow ? (
        <AddForm key={editingId} fields={fields} initialValues={editingRow}
          onSubmit={(vals) => { updateDocument(editingId, vals); setEditingId(null); }} submitLabel="Save Changes" />
      ) : canEdit && (
        <AddForm fields={fields} onSubmit={addDocument} submitLabel="+ Add Document" />
      )}
      <div className="text-xs mb-3" style={{ color: "var(--text-label)" }}>This records the document's details — actual file storage isn't wired up in this prototype yet, so no file bytes are uploaded.</div>
      <Table title="Documents" columns={["ID", "Document", "Module", "Type", "Uploaded By", "Date", ""]} rows={documents}
        renderRow={(d) => (<><Td mono>{d.id}</Td><Td>{d.name}</Td><Td>{d.module}</Td><Td>{d.type}</Td><Td>{d.uploadedBy}</Td><Td mono>{d.date}</Td>
          <Td>{isAdmin && <Button variant="ghost" onClick={() => setEditingId(d.id)}>Edit</Button>}{canEdit && <Button variant="danger" onClick={() => deleteDocument(d.id)}>Delete</Button>}</Td></>)} />
    </div>
  );
}

/* ================================================================
   Reports / Approvals / Users / Settings
   ================================================================ */

function ReportsPage() {
  const { leads, opportunities, salesOrders, orderTotal, materials, products, totalStock, employees, payroll, purchaseRequests, leaves, expenses, money } = useApp();
  const pipelineValue = opportunities.reduce((s, o) => s + o.value, 0);
  const salesThisMonth = salesOrders.filter((o) => o.date.startsWith("2026-08")).reduce((s, o) => s + orderTotal(o), 0);
  const lowStockCount = materials.filter((m) => m.stock < m.reorder).length + products.filter((p) => totalStock(p) < p.reorder).length;
  const pendingApprovals = purchaseRequests.filter((p) => p.status === "Pending").length + leaves.filter((l) => l.status === "Pending").length + expenses.filter((e) => e.status === "Pending").length;
  const payrollCost = payroll.reduce((s, p) => s + (p.base - p.deductions), 0);
  return (
    <div>
      <SectionHeader title="Reports" sub="Cross-module summary. Deeper report builders can be added as this system grows." />
      <div className="grid grid-cols-4 gap-4">
        <Card title="Open Leads" value={leads.length} />
        <Card title="Pipeline Value" value={money(pipelineValue)} accent="#3A5C86" />
        <Card title="Sales This Month" value={money(salesThisMonth)} accent="#3F7D5C" />
        <Card title="Low Stock Items" value={lowStockCount} accent="#A64B3A" />
        <Card title="Pending Approvals" value={pendingApprovals} accent="#8A6A2E" />
        <Card title="Headcount" value={employees.length} />
        <Card title="Payroll Cost (this run)" value={money(payrollCost)} accent="#6B4FA0" />
      </div>
    </div>
  );
}
function ApprovalsPage() {
  const { purchaseRequests, leaves, expenses, matById, approvePR, setLeaveStatus, setExpenseStatus, money, canEdit } = useApp();
  const items = [
    ...purchaseRequests.filter((p) => p.status === "Pending").map((p) => ({ type: "Purchase Request", id: p.id, detail: `${p.qty} × ${matById[p.item].name}`, requester: p.requestedBy, date: p.date, approve: () => approvePR(p.id, "Approved"), reject: () => approvePR(p.id, "Rejected") })),
    ...leaves.filter((l) => l.status === "Pending").map((l) => ({ type: "Leave Request", id: l.id, detail: `${l.type} leave, ${l.from} to ${l.to}`, requester: l.employee, date: l.from, approve: () => setLeaveStatus(l.id, "Approved"), reject: () => setLeaveStatus(l.id, "Rejected") })),
    ...expenses.filter((e) => e.status === "Pending").map((e) => ({ type: "Expense Claim", id: e.id, detail: `${e.category} — ${money(e.amount)}`, requester: e.employee, date: e.date, approve: () => setExpenseStatus(e.id, "Approved"), reject: () => setExpenseStatus(e.id, "Rejected") })),
  ];
  return (
    <div>
      <SectionHeader title="Approvals" sub="Everything across the system waiting on a decision." />
      <Table title="Approvals" columns={["Type", "Ref", "Detail", "Requester", "Date", ""]} rows={items}
        renderRow={(it) => (
          <>
            <Td><Pill>{it.type}</Pill></Td><Td mono>{it.id}</Td><Td>{it.detail}</Td><Td>{it.requester}</Td><Td mono>{it.date}</Td>
            <Td>{canEdit && <><Button onClick={it.approve}>Approve</Button><Button variant="danger" onClick={it.reject}>Reject</Button></>}</Td>
          </>
        )} />
      {items.length === 0 && <div className="text-sm mt-4" style={{ color: "var(--text-secondary)" }}>Nothing waiting on approval right now.</div>}
    </div>
  );
}
function UsersPage() {
  const { users, addUser, updateUserRole, deleteUser, canEdit, user: currentUser } = useApp();
  return (
    <div>
      <SectionHeader title="Users & Permissions" sub="Accounts and the modules each role can access." />
      <div className="text-sm font-semibold mb-2" style={{ color: "var(--heading)" }}>Users</div>
      {canEdit && <AddForm
        fields={[{ key: "name", label: "Name" }, { key: "username", label: "Username" }, { key: "role", label: "Role", type: "select", options: ROLES_LIST }]}
        onSubmit={addUser} submitLabel="+ Add User" />}
      <Table title="Users & Permissions" columns={["ID", "Name", "Username", "Role", ""]} rows={users}
        renderRow={(u) => (
          <>
            <Td mono>{u.id}</Td><Td>{u.name}</Td><Td mono>{u.username}</Td>
            <Td>
              {canEdit ? (
                <select value={u.role} onChange={(e) => updateUserRole(u.id, e.target.value)}
                  className="text-xs px-2 py-1 rounded-sm" style={{ border: "1px solid var(--border-strong)" }}>
                  {ROLES_LIST.map((r) => <option key={r} value={r}>{r}</option>)}
                </select>
              ) : <RolePill role={u.role} />}
            </Td>
            <Td>{canEdit && u.id !== currentUser?.id && <Button variant="danger" onClick={() => deleteUser(u.id)}>Delete</Button>}</Td>
          </>
        )} />
      <div className="text-sm font-semibold mt-6 mb-2" style={{ color: "var(--heading)" }}>Module access by role</div>
      <Table title="Users & Permissions" columns={["Module", ...ROLES_LIST]} rows={MODULES}
        renderRow={(m) => (<><Td>{m.label}</Td>{ROLES_LIST.map((r) => <Td key={r} mono>{MODULE_ACCESS[m.id].includes(r) ? "✓" : "—"}</Td>)}</>)} />
    </div>
  );
}
function SettingsPage() {
  const { currencies } = useApp();
  const [companyName, setCompanyName] = useState("CodePulse Manufacturing");
  const [baseCurrency, setBaseCurrency] = useState("USD");
  const [fiscalStart, setFiscalStart] = useState("January");
  const [defaultWarehouse, setDefaultWarehouse] = useState("BR-HQ");
  return (
    <div>
      <SectionHeader title="Settings" sub="Company-wide configuration (demo only — not persisted)." />
      <div className="p-4 rounded-md max-w-md space-y-4" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
        <div><div className="text-[11px] uppercase tracking-wide mb-1" style={{ color: "var(--text-label)" }}>Company Name</div>
          <input value={companyName} onChange={(e) => setCompanyName(e.target.value)} className="w-full text-sm px-2 py-1.5 rounded-sm" style={{ border: "1px solid var(--border-strong)" }} /></div>
        <div><div className="text-[11px] uppercase tracking-wide mb-1" style={{ color: "var(--text-label)" }}>Base Currency</div>
          <select value={baseCurrency} onChange={(e) => setBaseCurrency(e.target.value)} className="text-sm px-2 py-1.5 rounded-sm" style={{ border: "1px solid var(--border-strong)" }}>
            {currencies.map((c) => <option key={c.code} value={c.code}>{c.code}</option>)}
          </select></div>
        <div><div className="text-[11px] uppercase tracking-wide mb-1" style={{ color: "var(--text-label)" }}>Fiscal Year Start</div>
          <select value={fiscalStart} onChange={(e) => setFiscalStart(e.target.value)} className="text-sm px-2 py-1.5 rounded-sm" style={{ border: "1px solid var(--border-strong)" }}>
            {["January", "April", "July", "October"].map((m) => <option key={m} value={m}>{m}</option>)}
          </select></div>
        <div><div className="text-[11px] uppercase tracking-wide mb-1" style={{ color: "var(--text-label)" }}>Default Warehouse</div>
          <select value={defaultWarehouse} onChange={(e) => setDefaultWarehouse(e.target.value)} className="text-sm px-2 py-1.5 rounded-sm" style={{ border: "1px solid var(--border-strong)" }}>
            {BRANCHES.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}
          </select></div>
        <Button variant="accent" onClick={() => alert("Settings saved (demo only — resets when the session ends).")}>Save Settings</Button>
      </div>
    </div>
  );
}

function CurrenciesPage() {
  const { currencies, currency, updateCurrencyRate, addCurrency, canEdit } = useApp();
  const [edits, setEdits] = useState({});
  return (
    <div>
      <SectionHeader title="Currencies" sub="Exchange rates used everywhere money is shown or entered across the system. Rates are against USD." />
      {canEdit && <AddForm fields={[{ key: "code", label: "Code (e.g. CAD)", default: "" }, { key: "rate", label: "Rate vs USD", type: "number", default: 1.35 }]} onSubmit={addCurrency} submitLabel="+ Add Currency" />}
      <Table title="Currencies" columns={["Code", "Rate vs USD", "Currently Displaying", ""]} rows={currencies}
        renderRow={(c) => (
          <>
            <Td mono>{c.code}</Td>
            <Td mono>
              {canEdit ? (
                <input type="number" defaultValue={c.rate} onChange={(e) => setEdits((old) => ({ ...old, [c.code]: +e.target.value }))}
                  className="text-sm px-2 py-1 rounded-sm w-24" style={{ border: "1px solid var(--border-strong)" }} />
              ) : c.rate}
            </Td>
            <Td>{c.code === currency && <Pill>Active</Pill>}</Td>
            <Td>{canEdit && <Button onClick={() => updateCurrencyRate(c.code, edits[c.code] ?? c.rate)}>Update Rate</Button>}</Td>
          </>
        )} />
      <div className="text-sm mt-3" style={{ color: "var(--text-secondary)" }}>Any amount entered anywhere in the system (deals, expenses, budgets, etc.) can be typed in any of these currencies — it's converted and stored in USD automatically, then displayed back in whichever currency is selected in the top bar.</div>
    </div>
  );
}
