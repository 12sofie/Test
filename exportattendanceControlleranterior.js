const db = require("../config/db");
const dayjs = require("dayjs");
const ExcelJS = require("exceljs");
const utc = require("dayjs/plugin/utc");
const timezone = require("dayjs/plugin/timezone");
const isoWeek = require("dayjs/plugin/isoWeek");

// Extend dayjs with plugins
dayjs.extend(utc);
dayjs.extend(timezone);
dayjs.extend(isoWeek);

// Reporte de Día
// exports.exportAttendance = async (req, res) => {
//   try {
//     const { selectedDate } = req.body;
//     if (!selectedDate || typeof selectedDate !== "string") {
//       throw new Error("selectedDate is missing or invalid.");
//     }

//     const NO_HOUR_LABEL = "(sin hora)";
//     const NO_SHOW_LABEL = "(No vino, no marco asistencia)";
//     const LATE_THRESHOLD = "07:01:00"; // entradas estrictamente MAYORES a esto son tardanza

//     // colores
//     const HEADER_BAND = "E6F0FA";
//     const SUB_BAND = "F2F2F2";
//     const GREEN_E = "C6EFCE"; // header Entrada
//     const RED_S = "FFC7CE"; // header Salida
//     const YELLOW_D = "FFF2CC"; // header Despacho
//     const GRAY_P = "D9D9D9"; // header permisos
//     const BLUE_DIF = "BDD7EE"; // celda permiso diferido
//     const BLUE_SOL = "DDEBF7"; // celda permiso solicitado
//     const CELL_GREEN = "E2F0D9"; // celda entrada
//     const CELL_RED = "F8CBAD"; // celda salida
//     const CELL_YELLOW = "FFF2CC"; // celda despacho
//     const fmtTime = (t) =>
//       t ? dayjs(t, "HH:mm:ss").format("hh:mm:ss A") : null;

//     // Planilla
//     const [payrollRows] = await db.query(`
//       SELECT e.employeeID, COALESCE(pt.payrollName, '') AS payrollName
//       FROM employees_emp e
//       LEFT JOIN payrolltype_emp pt ON pt.payrollTypeID = e.payrollTypeID
//     `);
//     const payrollMap = new Map(
//       payrollRows.map((r) => [String(r.employeeID), r.payrollName])
//     );

//     // Empleados activos + excepción (FK directa a exceptions_emp)
//     const [activeEmployees] = await db.query(`
//       SELECT 
//         e.employeeID,
//         TRIM(CONCAT(e.firstName,' ',COALESCE(e.middleName,''),' ',e.lastName)) AS employeeName,
//         CASE WHEN UPPER(COALESCE(ex.exceptionName, '')) = 'LACTANCIA'  THEN 1 ELSE 0 END AS isLactation,
//         CASE WHEN UPPER(COALESCE(ex.exceptionName, '')) = 'EMBARAZADA' THEN 1 ELSE 0 END AS isPregnant
//       FROM employees_emp e
//       LEFT JOIN exceptions_emp ex ON ex.exceptionID = e.exceptionID
//       WHERE e.isActive = 1
//       ORDER BY e.employeeID
//     `);

//     // Asistencia del día
//     const [attendanceRows] = await db.query(
//       `
//       SELECT a.employeeID,
//              TIME_FORMAT(a.entryTime,'%H:%i:%s') AS entryTime,
//              TIME_FORMAT(a.exitTime, '%H:%i:%s') AS exitTime,
//              a.comment AS attendanceComment
//       FROM h_attendance_emp a
//       WHERE a.date = ?
//     `,
//       [selectedDate]
//     );
//     const attendanceByEmp = new Map(
//       attendanceRows.map((r) => [String(r.employeeID), r])
//     );

//     // Permisos del día (cronológico)
//     const [permRows] = await db.query(
//       `
//       SELECT p.permissionID, p.employeeID, p.request, p.comment,
//              TIME_FORMAT(p.exitPermission,  '%H:%i:%s') AS exitPermission,
//              TIME_FORMAT(p.entryPermission, '%H:%i:%s') AS entryPermission,
//              p.createdDate
//       FROM permissionattendance_emp p
//       WHERE p.date = ?
//         AND p.isApproved = 1
//       ORDER BY p.createdDate ASC, p.permissionID ASC
//     `,
//       [selectedDate]
//     );

//     const permsByEmp = new Map();
//     permRows.forEach((p) => {
//       const k = String(p.employeeID);
//       if (!permsByEmp.has(k)) permsByEmp.set(k, []);
//       permsByEmp.get(k).push(p);
//     });

//     // Despachos
//     const [dispatchRows] = await db.query(
//       `
//       SELECT d.employeeID,
//              TIME_FORMAT(d.exitTimeComplete,'%H:%i:%s') AS dispatchTime,
//              d.comment AS dispatchComment
//       FROM dispatching_emp d
//       WHERE d.date = ?
//     `,
//       [selectedDate]
//     );
//     const dispatchByEmp = new Map(
//       dispatchRows.map((d) => [String(d.employeeID), d])
//     );
//     const includeDispatch = dispatchRows.length > 0;

//     // Construir filas (hasta 5 permisos)
//     const rows = activeEmployees.map((emp) => {
//       const k = String(emp.employeeID);
//       const att = attendanceByEmp.get(k) || {};
//       const rawEntry = att.entryTime || null; // "HH:mm:ss" o null
//       const entryTime = fmtTime(att.entryTime) || "";
//       const exitTime = fmtTime(att.exitTime) || "";
//       const exitComment = att.attendanceComment || "";

//       const chronological = (permsByEmp.get(k) || []).slice(0, 5);

//       // Permisos formateados
//       const permSlots = chronological.map(p => {
//         const S = fmtTime(p.exitPermission) || "";
//         const R = fmtTime(p.entryPermission) || "";
//         const type = p.request === 0 ? "Dif" : "Sol";
//         const reason = p.comment || "";
//         return { S, R, type, reason, request: p.request };
//       });

//       // Si NO marcó asistencia y existe un permiso SOLICITADO sin horas, rotular como "No vino..."
//       const noAttendance = !entryTime && !exitTime;
//       if (noAttendance) {
//         const idxSol = chronological.findIndex(
//           (p) =>
//             p.request === 1 && // SOLICITADO
//             !p.exitPermission &&
//             !p.entryPermission // SIN horas de permiso
//         );
//         if (idxSol !== -1 && permSlots[idxSol]) {
//           permSlots[idxSol].S = NO_SHOW_LABEL;
//           permSlots[idxSol].R = NO_SHOW_LABEL;
//           permSlots[idxSol].type = "Sol";
//           permSlots[idxSol].__noShow = true;
//         }
//       }

//       const dsp = dispatchByEmp.get(k);
//       const dispatchTime = fmtTime(dsp?.dispatchTime) || "";
//       const dispatchComment = dsp?.dispatchComment ?? null;

//       // Etiquetas combinadas P/L/E/T
//       const hasP = (permsByEmp.get(k) || []).length > 0; // permiso aprobado ese día
//       const hasL = Number(emp.isLactation) === 1;
//       const hasE = Number(emp.isPregnant) === 1;

//       // Tardanza: sólo si tiene hora de entrada y es estrictamente mayor a 07:01:00
//       const hasT = !!rawEntry && rawEntry > LATE_THRESHOLD;

//       const tagParts = [];
//       if (hasP) tagParts.push("P");
//       if (hasL) tagParts.push("L");
//       if (hasE) tagParts.push("E");
//       if (hasT) tagParts.push("T");
//       const tags = tagParts.join(".");

//       return {
//         employeeID: emp.employeeID,
//         employeeName: emp.employeeName,
//         tags, // (P/L/E/T)
//         entryTime,
//         exitTime,
//         exitComment,
//         permSlots,
//         dispatchTime,
//         dispatchComment,
//         payrollName: payrollMap.get(k) || "",
//       };
//     });

//     const maxPermsInDay = rows.reduce(
//       (mx, r) => Math.max(mx, r.permSlots.length),
//       0
//     );

//     // Excel
//     const workbook = new ExcelJS.Workbook();
//     const ws = workbook.addWorksheet("Asistencia");

//     const headers = ["Item", "Código", "Empleado", "(P/L/E/T)", "E", "S"];
//     const widths = [10, 10, 30, 16, 14, 14];
//     for (let i = 1; i <= maxPermsInDay; i++) { headers.push(`P${i}S`, `P${i}R`); widths.push(12, 12); }
//     if (includeDispatch) { headers.push("D"); widths.push(12); }
//     headers.push("Tipo de Planilla", "Comentarios"); widths.push(16, 40);

//     // Mapa de columnas
//     const COL = {
//       ITEM: 1,
//       CODE: 2,
//       EMP: 3,
//       TAGS: 4,
//       E: 5,
//       S: 6,
//       P_START: 7,
//       getPCol: (j /*0-based*/) => 7 + j * 2,
//     };

//     // Título / subtítulo
//     const ncols = headers.length;
//     const title = ws.addRow([
//       `Reporte de Asistencia - Día ${dayjs(selectedDate).format("DD/MM/YYYY")}`,
//     ]);
//     ws.mergeCells(title.number, 1, title.number, ncols);
//     title.height = 28;
//     for (let c = 1; c <= ncols; c++) {
//       const cell = ws.getCell(title.number, c);
//       cell.font = { name: "Calibri", size: 16, bold: true };
//       cell.alignment = { horizontal: "center", vertical: "middle" };
//       cell.fill = {
//         type: "pattern",
//         pattern: "solid",
//         fgColor: { argb: HEADER_BAND },
//       };
//     }

//     const subtitle = ws.addRow([`Empleados activos: ${rows.length}  | P: Permisos / L: Lactancia / E: Embarazada / T: Tarde / EP: Permiso Especial `]);
//     ws.mergeCells(subtitle.number, 1, subtitle.number, ncols);
//     subtitle.height = 20;
//     for (let c = 1; c <= ncols; c++) {
//       const cell = ws.getCell(subtitle.number, c);
//       cell.font = { name: "Calibri", size: 12, bold: true };
//       cell.alignment = { horizontal: "center", vertical: "middle" };
//       cell.fill = {
//         type: "pattern",
//         pattern: "solid",
//         fgColor: { argb: SUB_BAND },
//       };
//     }

//     // Header row
//     const headerRow = ws.addRow(headers);
//     headerRow.eachCell((cell) => {
//       cell.font = { name: "Calibri", size: 12, bold: true };
//       cell.alignment = { horizontal: "center", vertical: "middle" };
//       cell.border = {
//         top: { style: "thin" },
//         bottom: { style: "thin" },
//         left: { style: "thin" },
//         right: { style: "thin" },
//       };
//       cell.fill = {
//         type: "pattern",
//         pattern: "solid",
//         fgColor: { argb: GRAY_P },
//       };
//     });
//     const colorHeader = (idx, argb) => {
//       headerRow.getCell(idx).fill = {
//         type: "pattern",
//         pattern: "solid",
//         fgColor: { argb },
//       };
//     };
//     colorHeader(COL.E, GREEN_E); // E
//     colorHeader(COL.S, RED_S); // S
//     if (includeDispatch) {
//       const dIndex = COL.S + maxPermsInDay * 2 + 1;
//       colorHeader(dIndex, YELLOW_D); // D
//     }

//     // Datos
//     rows.forEach((r, i) => {
//       // Item, Código, Empleado, Etiquetas, E, S
//       const data = [
//         i + 1,
//         r.employeeID,
//         r.employeeName,
//         r.tags || "",
//         r.entryTime,
//         r.exitTime,
//       ];

//       for (let j = 0; j < maxPermsInDay; j++) {
//         const p = r.permSlots[j];
//         if (!p) {
//           data.push("", "");
//           continue;
//         }

//         const label = (txt, isDif, reason) => {
//           if (txt) return txt;
//           if (isDif)
//             return reason
//               ? `${NO_HOUR_LABEL} – Justificación: ${reason}`
//               : NO_HOUR_LABEL;
//           return "";
//         };

//         // si venía tag __noShow => poner el rótulo en ambas
//         if (p.__noShow) {
//           data.push(NO_SHOW_LABEL, NO_SHOW_LABEL);
//         } else {
//           data.push(
//             label(p.S, p.type === "Dif", p.reason),
//             label(p.R, p.type === "Dif", p.reason)
//           );
//         }
//       }

//       if (includeDispatch) data.push(r.dispatchTime || "");

//       const comments = [];
//       r.permSlots.forEach((p, idx) => {
//         if (!p) return;
//         if (p.type === "Dif" && !p.S && !p.R) {
//           comments.push(
//             `P${idx + 1}: sin hora${p.reason ? ` (Justificación: ${p.reason})` : ""
//             }`
//           );
//         }
//         if (p.__noShow) {
//           comments.push(`P${idx + 1}: No vino, no marco asistencia`);
//         }
//       });
//       if (r.exitComment) comments.push(`Salida: ${r.exitComment}`);
//       if (includeDispatch) {
//         if (r.dispatchComment === 1 || r.dispatchComment === "1")
//           comments.push("Despacho: Cumplimiento de Meta");
//         else if (r.dispatchComment && r.dispatchComment !== "0")
//           comments.push(`Despacho: ${r.dispatchComment}`);
//       }

//       data.push(r.payrollName || "", comments.join(" | "));
//       const row = ws.addRow(data);

//       // estilos
//       row.eachCell((cell) => {
//         cell.font = { name: "Calibri", size: 11 };
//         cell.alignment = { horizontal: "left", vertical: "middle" };
//         cell.border = {
//           top: { style: "thin" },
//           bottom: { style: "thin" },
//           left: { style: "thin" },
//           right: { style: "thin" },
//         };
//       });

//       if (r.entryTime) row.getCell(COL.E).fill = { type: "pattern", pattern: "solid", fgColor: { argb: CELL_GREEN } };
//       if (r.exitTime) row.getCell(COL.S).fill = { type: "pattern", pattern: "solid", fgColor: { argb: CELL_RED } };

//       for (let j = 0; j < r.permSlots.length; j++) {
//         const base = COL.getPCol(j);
//         const p = r.permSlots[j];
//         if (!p) continue;
//         const argb = p.type === "Dif" ? BLUE_DIF : BLUE_SOL;
//         // sólo coloreo si hay texto visible
//         if (row.getCell(base).value) row.getCell(base).fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
//         if (row.getCell(base + 1).value) row.getCell(base + 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
//       }
//       if (includeDispatch) {
//         const dCol = COL.S + maxPermsInDay * 2 + 1;
//         if (r.dispatchTime)
//           row.getCell(dCol).fill = {
//             type: "pattern",
//             pattern: "solid",
//             fgColor: { argb: CELL_YELLOW },
//           };
//       }
//     });

//     // Altura y anchos
//     ws.eachRow((row, r) => {
//       if (r > 3) row.height = 20;
//     });
//     headers.forEach((_, i) => (ws.getColumn(i + 1).width = widths[i] || 12));

//     const buffer = await workbook.xlsx.writeBuffer();
//     const filename = `asistencia_dia_${dayjs(selectedDate).format(
//       "YYYYMMDD"
//     )}.xlsx`;
//     res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
//     res.setHeader(
//       "Content-Type",
//       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
//     );
//     res.send(buffer);
//   } catch (err) {
//     console.error(
//       "Error al exportar asistencia (día):",
//       err.stack || err.message
//     );
//     res
//       .status(500)
//       .send({
//         message: `Error interno del servidor al generar el archivo Excel: ${err.message}`,
//       });
//   }
// };

exports.exportAttendance = async (req, res) => {
  try {
    const { selectedDate } = req.body;
    if (!selectedDate || typeof selectedDate !== "string") {
      throw new Error("selectedDate is missing or invalid.");
    }

    // Semana (Lunes - Domingo) basada en selectedDate
    const sd = dayjs(selectedDate);
    const weekdayMon0 = (sd.day() + 6) % 7; // Lunes=0 ... Domingo=6
    const weekStart = sd.subtract(weekdayMon0, "day").format("YYYY-MM-DD");
    const weekEnd = sd
      .subtract(weekdayMon0, "day")
      .add(6, "day")
      .format("YYYY-MM-DD");

    const NO_HOUR_LABEL = "(sin hora)";
    const NO_SHOW_LABEL = "(No vino, no marco asistencia)";
    const LATE_THRESHOLD = "07:01:00"; // entradas estrictamente MAYORES a esto son tardanza

    // colores
    const HEADER_BAND = "E6F0FA";
    const SUB_BAND = "F2F2F2";
    const GREEN_E = "C6EFCE"; // header Entrada
    const RED_S = "FFC7CE"; // header Salida
    const YELLOW_D = "FFF2CC"; // header Despacho
    const GRAY_P = "D9D9D9"; // header permisos
    const BLUE_DIF = "BDD7EE"; // celda permiso diferido
    const BLUE_SOL = "DDEBF7"; // celda permiso solicitado
    const CELL_GREEN = "E2F0D9"; // celda entrada
    const CELL_RED = "F8CBAD"; // celda salida
    const CELL_YELLOW = "FFF2CC"; // celda despacho
    const DP_TAG_FILL = "F4CCCC"; // celda DP (despedido)
    const fmtTime = (t) => (t ? dayjs(t, "HH:mm:ss").format("hh:mm:ss A") : null);

    // Planilla
    const [payrollRows] = await db.query(`
      SELECT e.employeeID, COALESCE(pt.payrollName, '') AS payrollName
      FROM employees_emp e
      LEFT JOIN payrolltype_emp pt ON pt.payrollTypeID = e.payrollTypeID
    `);
    const payrollMap = new Map(payrollRows.map((r) => [String(r.employeeID), r.payrollName]));

    // Empleados activos + despedidos en la semana actual (para incluirlos en el reporte)
    const [activeEmployees] = await db.query(
      `
  SELECT 
    e.employeeID,
    TRIM(CONCAT(e.firstName,' ',COALESCE(e.middleName,''),' ',e.lastName)) AS employeeName,
    CASE WHEN UPPER(COALESCE(ex.exceptionName, '')) = 'LACTANCIA'  THEN 1 ELSE 0 END AS isLactation,
    CASE WHEN UPPER(COALESCE(ex.exceptionName, '')) = 'EMBARAZADA' THEN 1 ELSE 0 END AS isPregnant,
    CASE WHEN d.employeeID IS NOT NULL THEN 1 ELSE 0 END AS isDismissedWeek,
    d.dateDismissal
  FROM employees_emp e
  LEFT JOIN exceptions_emp ex ON ex.exceptionID = e.exceptionID
  LEFT JOIN (
      SELECT hd.employeeID, MAX(DATE(hd.dateDismissal)) AS dateDismissal
      FROM h_dismissal_emp hd
      WHERE DATE(hd.dateDismissal) BETWEEN ? AND ?
      GROUP BY hd.employeeID
  ) d ON d.employeeID = e.employeeID
  WHERE e.isActive = 1
     OR d.employeeID IS NOT NULL
  ORDER BY e.employeeID
  `,
      [weekStart, weekEnd]
    );

    // ========= DETECTAR INCAPACIDAD/MATERNIDAD (solo para selectedDate) =========
    const [incRows] = await db.query(
      `
      SELECT employeeID, kind
      FROM (
        SELECT
          d.employeeID,
          'DIS' AS kind
        FROM disability_emp d
        WHERE DATE(d.startDate) <= ? AND DATE(d.endDate) >= ?

        UNION ALL

        SELECT
          m.employeeID,
          'MAT' AS kind
        FROM maternity_emp m
        WHERE DATE(m.startDate) <= ? AND DATE(m.endDate) >= ?
      ) x
      `,
      [selectedDate, selectedDate, selectedDate, selectedDate]
    );

    // Map: employeeID -> kind ('DIS' or 'MAT')
    const incapByEmp = new Map();
    incRows.forEach((r) => {
      // si hubiera ambos por error, prioriza DIS sobre MAT
      const key = String(r.employeeID);
      const prev = incapByEmp.get(key);
      if (!prev) incapByEmp.set(key, r.kind);
      else if (prev !== "DIS" && r.kind === "DIS") incapByEmp.set(key, "DIS");
    });

    // Asistencia del día
    const [attendanceRows] = await db.query(
      `
      SELECT a.employeeID,
             TIME_FORMAT(a.entryTime,'%H:%i:%s') AS entryTime,
             TIME_FORMAT(a.exitTime, '%H:%i:%s') AS exitTime,
             a.comment AS attendanceComment
      FROM h_attendance_emp a
      WHERE a.date = ?
    `,
      [selectedDate]
    );
    const attendanceByEmp = new Map(attendanceRows.map((r) => [String(r.employeeID), r]));

    // Permisos del día (cronológico)
    const [permRows] = await db.query(
      `
      SELECT p.permissionID, p.employeeID, p.request, p.comment,
             TIME_FORMAT(p.exitPermission,  '%H:%i:%s') AS exitPermission,
             TIME_FORMAT(p.entryPermission, '%H:%i:%s') AS entryPermission,
             p.createdDate
      FROM permissionattendance_emp p
      WHERE p.date = ?
        AND p.isApproved = 1
      ORDER BY p.createdDate ASC, p.permissionID ASC
    `,
      [selectedDate]
    );

    const permsByEmp = new Map();
    permRows.forEach((p) => {
      const k = String(p.employeeID);
      if (!permsByEmp.has(k)) permsByEmp.set(k, []);
      permsByEmp.get(k).push(p);
    });

    // Despachos
    const [dispatchRows] = await db.query(
      `
      SELECT d.employeeID,
             TIME_FORMAT(d.exitTimeComplete,'%H:%i:%s') AS dispatchTime,
             d.comment AS dispatchComment
      FROM dispatching_emp d
      WHERE d.date = ?
    `,
      [selectedDate]
    );
    const dispatchByEmp = new Map(dispatchRows.map((d) => [String(d.employeeID), d]));
    const includeDispatch = dispatchRows.length > 0;

    // Construir filas (hasta 5 permisos)
    const rows = activeEmployees.map((emp) => {
      const k = String(emp.employeeID);

      // Si fue despedido y el selectedDate es posterior al despido: NO mostrar marcajes/permisos/despacho
      const dismissedDate = emp.dateDismissal ? dayjs(emp.dateDismissal).format("YYYY-MM-DD") : null;
      const isAfterDismissalDay =
        dismissedDate && dayjs(selectedDate).isAfter(dayjs(dismissedDate), "day"); // después del despido
      const isDismissalOrAfter =
        dismissedDate &&
        (dayjs(selectedDate).isAfter(dayjs(dismissedDate), "day") ||
          dayjs(selectedDate).isSame(dayjs(dismissedDate), "day")); // día del despido o después

      // asistencia/permisos/despacho mantienen la lógica ORIGINAL (no se bloquea por incapacidad)
      const att = isAfterDismissalDay ? {} : attendanceByEmp.get(k) || {};
      const rawEntry = att.entryTime || null; // "HH:mm:ss" o null
      const entryTime = fmtTime(att.entryTime) || "";
      const exitTime = fmtTime(att.exitTime) || "";
      const exitComment = att.attendanceComment || "";

      const chronological = isAfterDismissalDay ? [] : (permsByEmp.get(k) || []).slice(0, 5);

      // Permisos formateados
      const permSlots = chronological.map((p) => {
        const S = fmtTime(p.exitPermission) || "";
        const R = fmtTime(p.entryPermission) || "";
        const type = p.request === 0 ? "Dif" : "Sol";
        const reason = p.comment || "";
        return { S, R, type, reason, request: p.request };
      });

      // Si NO marcó asistencia y existe un permiso SOLICITADO sin horas, rotular como "No vino..."
      const noAttendance = !entryTime && !exitTime;
      if (noAttendance) {
        const idxSol = chronological.findIndex(
          (p) => p.request === 1 && !p.exitPermission && !p.entryPermission
        );
        if (idxSol !== -1 && permSlots[idxSol]) {
          permSlots[idxSol].S = NO_SHOW_LABEL;
          permSlots[idxSol].R = NO_SHOW_LABEL;
          permSlots[idxSol].type = "Sol";
          permSlots[idxSol].__noShow = true;
        }
      }

      const dsp = isAfterDismissalDay ? null : dispatchByEmp.get(k);
      const dispatchTime = fmtTime(dsp?.dispatchTime) || "";
      const dispatchComment = dsp?.dispatchComment ?? null;

      // ========= ESTADOS (PALABRAS) + DETECTAR INCAPACIDAD =========
      const hasDP = Number(emp.isDismissedWeek) === 1 && isDismissalOrAfter;
      const hasP = (permsByEmp.get(k) || []).length > 0 && !isAfterDismissalDay;
      const hasL = Number(emp.isLactation) === 1;
      const hasE = Number(emp.isPregnant) === 1;
      const hasT = !!rawEntry && rawEntry > LATE_THRESHOLD;

      // NUEVO: detectar incapacidad/maternidad hoy
      const incKind = incapByEmp.get(k); // 'DIS' | 'MAT' | undefined
      const hasINC = incKind === "DIS";
      const hasMAT = incKind === "MAT";

      const tagParts = [];
      if (hasDP) tagParts.push("DESPEDIDO");
      if (hasINC) tagParts.push("INCAPACIDAD");
      if (hasMAT) tagParts.push("MATERNIDAD");
      if (hasP) tagParts.push("PERMISO");
      if (hasL) tagParts.push("LACTANCIA");
      if (hasE) tagParts.push("EMBARAZADA");
      if (hasT) tagParts.push("TARDE");

      const tags = tagParts.join(" | ");

      return {
        employeeID: emp.employeeID,
        employeeName: emp.employeeName,
        tags, // ahora palabras
        entryTime,
        exitTime,
        exitComment,
        permSlots,
        dispatchTime,
        dispatchComment,
        payrollName: payrollMap.get(k) || "",
      };
    });

    const maxPermsInDay = rows.reduce((mx, r) => Math.max(mx, r.permSlots.length), 0);

    // Excel
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("Asistencia");

    // Header en palabras
    const headers = ["Item", "Código", "Empleado", "Estado", "E", "S"];
    const widths = [10, 10, 30, 26, 14, 14];

    for (let i = 1; i <= maxPermsInDay; i++) {
      headers.push(`P${i}S`, `P${i}R`);
      widths.push(12, 12);
    }
    if (includeDispatch) {
      headers.push("D");
      widths.push(12);
    }
    headers.push("Tipo de Planilla", "Comentarios");
    widths.push(16, 40);

    // Mapa de columnas
    const COL = {
      ITEM: 1,
      CODE: 2,
      EMP: 3,
      TAGS: 4,
      E: 5,
      S: 6,
      P_START: 7,
      getPCol: (j /*0-based*/) => 7 + j * 2,
    };

    // Título / subtítulo
    const ncols = headers.length;
    const title = ws.addRow([`Reporte de Asistencia - Día ${dayjs(selectedDate).format("DD/MM/YYYY")}`]);
    ws.mergeCells(title.number, 1, title.number, ncols);
    title.height = 28;
    for (let c = 1; c <= ncols; c++) {
      const cell = ws.getCell(title.number, c);
      cell.font = { name: "Calibri", size: 16, bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEADER_BAND } };
    }

    const subtitle = ws.addRow([
      `Empleados activos: ${rows.length}  | PERMISO / LACTANCIA / EMBARAZADA / TARDE / DESPEDIDO / INCAPACIDAD / MATERNIDAD`,
    ]);
    ws.mergeCells(subtitle.number, 1, subtitle.number, ncols);
    subtitle.height = 20;
    for (let c = 1; c <= ncols; c++) {
      const cell = ws.getCell(subtitle.number, c);
      cell.font = { name: "Calibri", size: 12, bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: SUB_BAND } };
    }

    // Header row
    const headerRow = ws.addRow(headers);
    headerRow.eachCell((cell) => {
      cell.font = { name: "Calibri", size: 12, bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
      };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRAY_P } };
    });

    const colorHeader = (idx, argb) => {
      headerRow.getCell(idx).fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
    };
    colorHeader(COL.E, GREEN_E); // E
    colorHeader(COL.S, RED_S); // S
    if (includeDispatch) {
      const dIndex = COL.S + maxPermsInDay * 2 + 1;
      colorHeader(dIndex, YELLOW_D); // D
    }

    // Datos
    rows.forEach((r, i) => {
      // Item, Código, Empleado, Etiquetas, E, S
      const data = [i + 1, r.employeeID, r.employeeName, r.tags || "", r.entryTime, r.exitTime];

      for (let j = 0; j < maxPermsInDay; j++) {
        const p = r.permSlots[j];
        if (!p) {
          data.push("", "");
          continue;
        }

        const label = (txt, isDif, reason) => {
          if (txt) return txt;
          if (isDif) return reason ? `${NO_HOUR_LABEL} – Justificación: ${reason}` : NO_HOUR_LABEL;
          return "";
        };

        // si venía tag __noShow => poner el rótulo en ambas
        if (p.__noShow) {
          data.push(NO_SHOW_LABEL, NO_SHOW_LABEL);
        } else {
          data.push(label(p.S, p.type === "Dif", p.reason), label(p.R, p.type === "Dif", p.reason));
        }
      }

      if (includeDispatch) data.push(r.dispatchTime || "");

      const comments = [];
      r.permSlots.forEach((p, idx) => {
        if (!p) return;
        if (p.type === "Dif" && !p.S && !p.R) {
          comments.push(`P${idx + 1}: sin hora${p.reason ? ` (Justificación: ${p.reason})` : ""}`);
        }
        if (p.__noShow) {
          comments.push(`P${idx + 1}: No vino, no marco asistencia`);
        }
      });
      if (r.exitComment) comments.push(`Salida: ${r.exitComment}`);
      if (includeDispatch) {
        if (r.dispatchComment === 1 || r.dispatchComment === "1") comments.push("Despacho: Cumplimiento de Meta");
        else if (r.dispatchComment && r.dispatchComment !== "0") comments.push(`Despacho: ${r.dispatchComment}`);
      }

      data.push(r.payrollName || "", comments.join(" | "));
      const row = ws.addRow(data);

      // estilos
      row.eachCell((cell) => {
        cell.font = { name: "Calibri", size: 11 };
        cell.alignment = { horizontal: "left", vertical: "middle" };
        cell.border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" },
        };
      });

      // Distintivo DESPEDIDO en la columna de etiquetas (antes DP)
      if ((r.tags || "").includes("DESPEDIDO")) {
        row.getCell(COL.TAGS).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: DP_TAG_FILL },
        };
        row.getCell(COL.TAGS).font = { name: "Calibri", size: 11, bold: true };
      }

      if (r.entryTime)
        row.getCell(COL.E).fill = { type: "pattern", pattern: "solid", fgColor: { argb: CELL_GREEN } };
      if (r.exitTime)
        row.getCell(COL.S).fill = { type: "pattern", pattern: "solid", fgColor: { argb: CELL_RED } };

      for (let j = 0; j < r.permSlots.length; j++) {
        const base = COL.getPCol(j);
        const p = r.permSlots[j];
        if (!p) continue;
        const argb = p.type === "Dif" ? BLUE_DIF : BLUE_SOL;
        if (row.getCell(base).value)
          row.getCell(base).fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
        if (row.getCell(base + 1).value)
          row.getCell(base + 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
      }

      if (includeDispatch) {
        const dCol = COL.S + maxPermsInDay * 2 + 1;
        if (r.dispatchTime)
          row.getCell(dCol).fill = { type: "pattern", pattern: "solid", fgColor: { argb: CELL_YELLOW } };
      }
    });

    // Altura y anchos
    ws.eachRow((row, r) => {
      if (r > 3) row.height = 20;
    });
    headers.forEach((_, i) => (ws.getColumn(i + 1).width = widths[i] || 12));

    const buffer = await workbook.xlsx.writeBuffer();
    const filename = `asistencia_dia_${dayjs(selectedDate).format("YYYYMMDD")}.xlsx`;
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.send(buffer);
  } catch (err) {
    console.error("Error al exportar asistencia (día):", err.stack || err.message);
    res.status(500).send({
      message: `Error interno del servidor al generar el archivo Excel: ${err.message}`,
    });
  }
};

// Reporte Semanal

// exports.exportWeeklyAttendance = async (req, res) => {
//   try {
//     const { weeklyAttendance, selectedMonth, selectedWeek, isActive } = req.body;

//     if (!weeklyAttendance || !Array.isArray(weeklyAttendance))
//       throw new Error("weeklyAttendance is missing or invalid.");
//     if (!selectedMonth || typeof selectedMonth !== "string")
//       throw new Error("selectedMonth is missing or invalid.");
//     if (!selectedWeek || typeof selectedWeek !== "string")
//       throw new Error("selectedWeek is missing or invalid.");

//     // ====== Helpers de tiempo ======
//     const parseTime = (t) =>
//       dayjs(
//         t,
//         ["HH:mm:ss", "H:mm:ss", "hh:mm:ss A", "h:mm:ss A"],
//         true
//       );

//     const fmt12smart = (t, { bias = "auto", refEntry } = {}) => {
//       if (!t) return "-";
//       let d = dayjs(
//         t,
//         [
//           "HH:mm:ss",
//           "H:mm:ss",
//           "hh:mm:ss A",
//           "h:mm:ss A",
//           "YYYY-MM-DD HH:mm:ss",
//           "YYYY-MM-DD hh:mm:ss A",
//         ],
//         true
//       );
//       if (!d.isValid()) return String(t);

//       const raw = String(t);
//       const hasMeridiem = /AM|PM/i.test(raw);
//       if (!hasMeridiem && bias === "exit" && refEntry) {
//         const e = dayjs(
//           refEntry,
//           ["HH:mm:ss", "H:mm:ss", "hh:mm:ss A", "h:mm:ss A"],
//           true
//         );
//         if (e.isValid()) {
//           const hrEntry = e.hour();
//           const hrExit = d.hour();
//           if (hrEntry >= 5 && hrEntry <= 9 && hrExit >= 1 && hrExit <= 6)
//             d = d.add(12, "hour");
//         }
//       }
//       return d.format("hh:mm:ss A");
//     };

//     const timeToMinutes = (t) => {
//       if (!t) return null;
//       const d = parseTime(t);
//       if (!d.isValid()) return null;
//       return d.hour() * 60 + d.minute();
//     };

//     const mergeIntervals = (intervals) => {
//       if (!intervals.length) return [];
//       const sorted = intervals
//         .slice()
//         .sort((a, b) => a[0] - b[0] || a[1] - b[1]);
//       const result = [];
//       let [curStart, curEnd] = sorted[0];
//       for (let i = 1; i < sorted.length; i++) {
//         const [s, e] = sorted[i];
//         if (s <= curEnd) {
//           if (e > curEnd) curEnd = e;
//         } else {
//           result.push([curStart, curEnd]);
//           curStart = s;
//           curEnd = e;
//         }
//       }
//       result.push([curStart, curEnd]);
//       return result;
//     };

//     const totalMinutes = (intervals) => {
//       const merged = mergeIntervals(intervals);
//       return merged.reduce((sum, [s, e]) => sum + (e - s), 0);
//     };

//     const coverageInSegment = (intervals, segStart, segEnd) => {
//       if (segStart == null || segEnd == null || segEnd <= segStart) return 0;
//       const clipped = [];
//       for (const [s, e] of intervals) {
//         const cs = Math.max(segStart, s);
//         const ce = Math.min(segEnd, e);
//         if (ce > cs) clipped.push([cs, ce]);
//       }
//       return totalMinutes(clipped);
//     };

//     // ====== Constantes / colores ======
//     const TITLE_BG = "1F3864";
//     const WHITE = "FFFFFF";
//     const GREEN_E = "C6EFCE";
//     const RED_S = "FFC7CE";
//     const GRAY_P = "D6DCE4";
//     const YELLOW_D = "FFEB9C";
//     const NO_HOUR = "(sin hora)";
//     const NO_SHOW = "(No vino, no marco asistencia)";

//     // Almuerzo genérico
//     const LUNCH_HOURS_DEFAULT = 0.75; // 45 minutos
//     const LUNCH_OFFSET_MIN = 5.75 * 60; // 5h45 después de la entrada

//     const dayHeaderColors = {
//       Lunes: "FFFFFF",
//       Martes: "FFC0CB",
//       Miércoles: "FFFFFF",
//       Jueves: "FFFFFF",
//       Viernes: "FFFFFF",
//       Sábado: "D3D3D3",
//       Domingo: "D3D3D3",
//     };

//     // ====== Empleados según isActive ======
//     const [activeEmployees] = await db.query(`
//       SELECT 
//         e.employeeID,
//         e.shiftID,
//         CONCAT(e.firstName,' ',COALESCE(e.middleName,''),' ',e.lastName) AS employeeName
//       FROM employees_emp e
//       WHERE e.isActive = ${isActive}
//       ORDER BY e.employeeID
//     `);
//     const activeSet = new Set(activeEmployees.map((e) => String(e.employeeID)));

//     // ====== Mapa de Tipo de Planilla ======
//     const [payrollRows] = await db.query(`
//       SELECT e.employeeID, COALESCE(pt.payrollName,'') AS payrollName
//       FROM employees_emp e
//       LEFT JOIN payrolltype_emp pt ON pt.payrollTypeID = e.payrollTypeID
//     `);
//     const payrollMap = new Map(
//       payrollRows.map((r) => [String(r.employeeID), r.payrollName || ""])
//     );

//     // ====== Detalle de turnos por día (detailshift_emp) ======
//     const [shiftDetailRows] = await db.query(`
//       SELECT shiftID, day, startTime, endTime
//       FROM detailsshift_emp
//     `);

//     const dayNameToIso = {
//       Monday: 1,
//       Tuesday: 2,
//       Wednesday: 3,
//       Thursday: 4,
//       Friday: 5,
//       Saturday: 6,
//       Sunday: 7,
//     };

//     // Mapa: `${shiftID}|${isoDay}` -> { startMin, endMin }
//     const shiftDetailMap = new Map();
//     for (const r of shiftDetailRows) {
//       const isoDay = dayNameToIso[r.day];
//       if (!isoDay) continue;
//       let s = timeToMinutes(r.startTime);
//       let e = timeToMinutes(r.endTime);
//       if (s == null || e == null) continue;
//       // Si la hora de salida es <= entrada, entonces cruza medianoche
//       if (e <= s) e += 24 * 60;
//       shiftDetailMap.set(`${r.shiftID}|${isoDay}`, { startMin: s, endMin: e });
//     }

//     // Devuelve la configuración de jornada según el turno y el día ISO (1–7)
//     const getShiftConfig = (shiftId, isoDay) => {
//       const key = `${shiftId}|${isoDay}`;
//       const base = shiftDetailMap.get(key);

//       if (base) {
//         const { startMin, endMin } = base;
//         const lunchHours = LUNCH_HOURS_DEFAULT;
//         const hrsLaborables = (endMin - startMin) / 60 - lunchHours;
//         const lunchThresholdMin = startMin + LUNCH_OFFSET_MIN;
//         return { startMin, endMin, lunchHours, lunchThresholdMin, hrsLaborables };
//       }

//       // Fallback: si no hay detalle, usamos un turno genérico 07:00–16:45
//       const startMin = 7 * 60;
//       const endMin = 16 * 60 + 45;
//       const lunchHours = LUNCH_HOURS_DEFAULT;
//       const hrsLaborables = (endMin - startMin) / 60 - lunchHours;
//       const lunchThresholdMin = startMin + LUNCH_OFFSET_MIN;
//       return { startMin, endMin, lunchHours, lunchThresholdMin, hrsLaborables };
//     };

//     // ====== Mapa de exceptionTime (permiso especial en horas) ======
//     const [excRows] = await db.query(`
//       SELECT e.employeeID, COALESCE(ex.exceptionTime, 0) AS exceptionTime
//       FROM employees_emp e
//       LEFT JOIN exceptions_emp ex ON ex.exceptionID = e.exceptionID
//       WHERE e.employeeID IN (${[...activeSet]
//         .map((id) => db.escape(id))
//         .join(",")})
//     `);
//     const exceptionTimeMap = new Map(
//       excRows.map((r) => [String(r.employeeID), Number(r.exceptionTime || 0)])
//     );

//     // ====== Semana ISO ======
//     let startOfWeek;
//     if (weeklyAttendance.length > 0) {
//       const firstDate = weeklyAttendance.map((r) => String(r.date)).sort()[0];
//       startOfWeek = dayjs(firstDate).startOf("isoWeek");
//     } else {
//       const yearNow = dayjs().year();
//       startOfWeek = dayjs()
//         .year(yearNow)
//         .isoWeek(parseInt(selectedWeek))
//         .startOf("isoWeek");
//     }
//     const endOfWeek = startOfWeek.add(6, "day");
//     const days = Array.from({ length: 7 }, (_, i) => {
//       const d = startOfWeek.add(i, "day");
//       return {
//         idx: i,
//         iso: d.isoWeekday(), // 1 = lunes ... 7 = domingo
//         date: d.format("YYYY-MM-DD"),
//         name:
//           d.format("dddd").charAt(0).toUpperCase() +
//           d.format("dddd").slice(1),
//         short: d.format("DD/MM"),
//         abbr: d.format("ddd"),
//       };
//     });

//     // ====== Indexar asistencias recibidas (entrada/salida/ despacho) ======
//     const attMap = new Map(
//       weeklyAttendance
//         .filter((r) => activeSet.has(String(r.employeeID)))
//         .map((r) => {
//           const key = `${String(r.employeeID)}|${String(r.date)}`;
//           return [
//             key,
//             {
//               entryTime: String(r.entryTime ?? ""),
//               exitTime: String(r.exitTime ?? ""),
//               dispatchingTime: String(r.dispatchingTime ?? ""),
//               dispatchingComment: String(r.dispatchingComment ?? ""),
//               exitComment: String(r.exitComment ?? ""),
//             },
//           ];
//         })
//     );

//     // ====== Grid base: todos los empleados × 7 días ======
//     const san = [];
//     for (const emp of activeEmployees) {
//       const empId = String(emp.employeeID);
//       for (const d of days) {
//         const key = `${empId}|${d.date}`;
//         const rec =
//           attMap.get(key) || {
//             entryTime: "",
//             exitTime: "",
//             dispatchingTime: "",
//             dispatchingComment: "",
//             exitComment: "",
//           };
//         san.push({
//           employeeID: empId,
//           employeeName: String(emp.employeeName ?? ""),
//           date: d.date,
//           ...rec,
//         });
//       }
//     }

//     // ====== Permisos (aprobados) de la semana ======
//     const [permWeekRows] = await db.query(
//       `
//       SELECT employeeID,
//              DATE(date) AS date,
//              TIME_FORMAT(exitPermission,  '%H:%i:%s') AS exitFmt,
//              TIME_FORMAT(entryPermission, '%H:%i:%s') AS entryFmt,
//              request,
//              comment,
//              lunchTime,
//              createdDate,
//              permissionID
//       FROM permissionattendance_emp
//       WHERE date BETWEEN ? AND ?
//         AND isApproved = 1
//       ORDER BY employeeID, date, createdDate, permissionID
//       `,
//       [startOfWeek.format("YYYY-MM-DD"), endOfWeek.format("YYYY-MM-DD")]
//     );

//     const permsByEmpDate = new Map();
//     for (const p of permWeekRows) {
//       const emp = String(p.employeeID);
//       if (!activeSet.has(emp)) continue;
//       const key = `${emp}|${dayjs(p.date).format("YYYY-MM-DD")}`;
//       if (!permsByEmpDate.has(key)) permsByEmpDate.set(key, []);
//       permsByEmpDate.get(key).push({
//         request: p.request == 0 || p.request === "0" ? 0 : 1,
//         exitFmt: p.exitFmt || "",
//         entryFmt: p.entryFmt || "",
//         comment: p.comment || "",
//         lunchTime: p.lunchTime,
//       });
//     }

//     // ====== Excepciones por día (L, E, EP) ======
//     let hasBridge = false;
//     let bridgeRows = [];
//     try {
//       const [rowsBridge] = await db.query(`
//         SELECT 
//           ee.employeeID,
//           ee.isActive,
//           ee.startDate,
//           ee.endDate,
//           UPPER(ex.exceptionName) AS exceptionName
//         FROM employee_exceptions ee
//         JOIN exceptions_emp ex ON ex.exceptionID = ee.exceptionID
//         WHERE ee.employeeID IN (${[...activeSet]
//           .map((id) => db.escape(id))
//           .join(",")})
//       `);
//       bridgeRows = rowsBridge || [];
//       hasBridge = bridgeRows.length > 0;
//     } catch (_) {
//       hasBridge = false;
//       bridgeRows = [];
//     }

//     const flagsByEmpDate = new Map();
//     if (hasBridge) {
//       for (const empId of activeSet) {
//         const myExc = bridgeRows.filter(
//           (r) => String(r.employeeID) === String(empId)
//         );
//         for (const d of days) {
//           const applies = (rec) => {
//             const active =
//               rec.isActive == null ||
//               rec.isActive === 1 ||
//               rec.isActive === "1";
//             const sdOk =
//               !rec.startDate ||
//               dayjs(d.date).isSameOrAfter(dayjs(rec.startDate), "day");
//             const edOk =
//               !rec.endDate ||
//               dayjs(d.date).isSameOrBefore(dayjs(rec.endDate), "day");
//             return active && sdOk && edOk;
//           };
//           const norm = (x) => (x || "").toString().trim().toUpperCase();
//           const L = myExc.some(
//             (r) => applies(r) && norm(r.exceptionName) === "LACTANCIA"
//           )
//             ? 1
//             : 0;
//           const E = myExc.some(
//             (r) => applies(r) && norm(r.exceptionName) === "EMBARAZADA"
//           )
//             ? 1
//             : 0;
//           const EP = myExc.some(
//             (r) => applies(r) && norm(r.exceptionName) === "PERMISO ESPECIAL"
//           )
//             ? 1
//             : 0;
//           flagsByEmpDate.set(`${empId}|${d.date}`, { L, E, EP });
//         }
//       }
//     } else {
//       const [fkRows] = await db.query(`
//         SELECT 
//           e.employeeID,
//           UPPER(COALESCE(ex.exceptionName,'')) AS exName
//         FROM employees_emp e
//         LEFT JOIN exceptions_emp ex ON ex.exceptionID = e.exceptionID
//         WHERE e.employeeID IN (${[...activeSet]
//           .map((id) => db.escape(id))
//           .join(",")})
//       `);
//       const fkMap = new Map(
//         fkRows.map((r) => [String(r.employeeID), r.exName || ""])
//       );
//       for (const empId of activeSet) {
//         const exName = (fkMap.get(String(empId)) || "").toUpperCase();
//         const L = exName === "LACTANCIA" ? 1 : 0;
//         const E = exName === "EMBARAZADA" ? 1 : 0;
//         const EP = 0;
//         for (const d of days)
//           flagsByEmpDate.set(`${empId}|${d.date}`, { L, E, EP });
//       }
//     }

//     // ====== Función para cálculo diario (H.T / H.A) ======
//     const computeDaySummary = (
//       entry,
//       exit,
//       permsList,
//       exceptionTime = 0,
//       shiftCfg
//     ) => {
//       const { startMin, endMin, lunchHours, lunchThresholdMin, hrsLaborables } =
//         shiftCfg;

//       const crossesMidnight = endMin > 24 * 60;

//       // Ajusta horas a la línea de tiempo del turno
//       const adjustForShift = (min) => {
//         if (min == null) return null;

//         if (crossesMidnight && min < startMin && min <= 12 * 60) {
//           // Turno que cruza medianoche:
//           // horas entre 00:00 y 12:00 pertenecen al "día siguiente"
//           min += 24 * 60;
//         }
//         return min;
//       };

//       // Minutos ya ajustados al turno
//       let entryMin = entry ? adjustForShift(timeToMinutes(entry)) : null;
//       let exitMin = exit ? adjustForShift(timeToMinutes(exit)) : null;

//       const entryMinClamped =
//         entryMin == null
//           ? null
//           : Math.min(Math.max(entryMin, startMin), endMin);
//       const exitMinClamped =
//         exitMin == null ? null : Math.min(Math.max(exitMin, startMin), endMin);

//       // ===== Permisos en bruto =====
//       const permisoIntervalsRaw = [];
//       let permisoLunchFlag = false;

//       for (const p of permsList) {
//         let s = adjustForShift(timeToMinutes(p.exitFmt));
//         let e = adjustForShift(timeToMinutes(p.entryFmt));
//         if (s == null || e == null) continue;
//         if (e <= s) continue;
//         permisoIntervalsRaw.push([s, e]);
//         if (Number(p.lunchTime || 0) === 1) permisoLunchFlag = true;
//       }

//       const permisoMinInShift = coverageInSegment(
//         permisoIntervalsRaw,
//         startMin,
//         endMin
//       );
//       let hrsPermiso = permisoMinInShift / 60.0;
//       if (permisoLunchFlag && hrsPermiso > 0) {
//         hrsPermiso = Math.max(hrsPermiso - lunchHours, 0);
//       }

//       const anyLunchFlag = permisoLunchFlag;

//       const soloEntradaSinSalida = !!entry && !exit;
//       const soloSalidaSinEntrada = !entry && !!exit;
//       const caso3 = soloEntradaSinSalida || soloSalidaSinEntrada;

//       let hrsAusente_entrada = 0;
//       let hrsAusente_salida = 0;
//       let hrsAusencia = 0;
//       let hrsTrabajadas = 0;

//       if (caso3) {
//         // Sólo entrada o sólo salida → no marcamos ausencia
//         hrsTrabajadas = 0;
//         hrsAusencia = 0;
//       } else {
//         const noEntry = !entry;
//         const noExit = !exit;

//         // Sin entrada, sin salida, pero CON permisos → ausencia parcial
//         if (noEntry && noExit && permisoMinInShift === 0) {
//           const netPermiso = hrsPermiso;
//           hrsAusencia = Math.max(hrsLaborables - netPermiso, 0);
//         } else {
//           // ---- GAP DE ENTRADA ----
//           if (entryMinClamped != null && entryMinClamped > startMin) {
//             const totalGapMin = entryMinClamped - startMin;
//             const coveredMin = coverageInSegment(
//               permisoIntervalsRaw,
//               startMin,
//               entryMinClamped
//             );
//             let baseMin = totalGapMin - coveredMin;
//             if (baseMin < 0) baseMin = 0;

//             // Si llega después del umbral de almuerzo y no hay permiso que marque almuerzo,
//             // descontamos el almuerzo del "castigo"
//             if (
//               baseMin > 0 &&
//               entryMinClamped >= lunchThresholdMin &&
//               !anyLunchFlag
//             ) {
//               baseMin -= lunchHours * 60;
//               if (baseMin < 0) baseMin = 0;
//             }

//             hrsAusente_entrada = baseMin / 60.0;
//           }

//           // ---- GAP DE SALIDA ----
//           if (exitMinClamped != null && exitMinClamped < endMin) {
//             const totalGapMin = endMin - exitMinClamped;
//             const coveredAfterExit = coverageInSegment(
//               permisoIntervalsRaw,
//               exitMinClamped,
//               endMin
//             );
//             let effectiveGapMin = totalGapMin - coveredAfterExit;
//             if (effectiveGapMin < 0) effectiveGapMin = 0;

//             let baseHours =
//               effectiveGapMin / 60.0 - Number(exceptionTime || 0);
//             if (baseHours < 0) baseHours = 0;

//             hrsAusente_salida = baseHours;
//           }
//           hrsAusencia = hrsAusente_entrada + hrsAusente_salida;
//         }

//         hrsTrabajadas = Math.max(
//           hrsLaborables - (hrsPermiso + hrsAusencia),
//           0
//         );
//       }

//       return {
//         hrsTrabajadas,
//         hrsAusencia,
//       };
//     };


//     // ====== Máximo de permisos por día ======
//     const maxPermsByDay = Array(7).fill(0);
//     for (const [key, list] of permsByEmpDate.entries()) {
//       const [, dateStr] = key.split("|");
//       const idx = dayjs(dateStr).diff(startOfWeek, "day");
//       if (idx < 0 || idx > 6) continue;
//       maxPermsByDay[idx] = Math.max(
//         maxPermsByDay[idx],
//         Math.min(list.length, 5)
//       );
//     }

//     // ====== ¿Hay despacho por día? ======
//     const includeDispatchByDay = Array(7).fill(false);
//     for (let i = 0; i < 7; i++) {
//       includeDispatchByDay[i] = san.some(
//         (r) =>
//           r.date === days[i].date &&
//           r.dispatchingTime &&
//           activeSet.has(r.employeeID)
//       );
//     }

//     // ====== Si no hay empleados ======
//     if (activeEmployees.length === 0) {
//       const workbook0 = new ExcelJS.Workbook();
//       const w = workbook0.addWorksheet("Asistencia Semanal");
//       w.addRow([
//         `Reporte Semanal - ${isActive == 0 ? "INACTIVOS" : "ACTIVOS"
//         } - (sin empleados)`,
//       ]);
//       w.mergeCells(1, 1, 1, 5);
//       w.getRow(1).font = { bold: true };
//       w.addRow([
//         "Item",
//         "Código",
//         "Empleado",
//         "Tipo de Planilla",
//         "Comentarios",
//       ]);
//       w.columns = [
//         { width: 10 },
//         { width: 10 },
//         { width: 30 },
//         { width: 18 },
//         { width: 60 },
//       ];
//       const buf0 = await workbook0.xlsx.writeBuffer();
//       const filename0 = `asistencia_semanal_${isActive == 0 ? "inactivos" : "activos"
//         }_semana${selectedWeek}_${dayjs().year()}.xlsx`;
//       res.setHeader(
//         "Content-Disposition",
//         `attachment; filename="${filename0}"`
//       );
//       res.setHeader(
//         "Content-Type",
//         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
//       );
//       return res.send(buf0);
//     }

//     // ====== Excel (estructura con H.T / H.A + permisos + despacho) ======
//     const workbook = new ExcelJS.Workbook();
//     const ws = workbook.addWorksheet("Asistencia Semanal");

//     const year = dayjs().year();
//     const monthName = dayjs()
//       .month(parseInt(selectedMonth))
//       .format("MMMM")
//       .toUpperCase();

//     // Por día:
//     const columnsPerDay = days.map(
//       (_, i) =>
//         1 + 4 + maxPermsByDay[i] * 2 + (includeDispatchByDay[i] ? 1 : 0)
//     );
//     const totalDataCols = 4 + columnsPerDay.reduce((a, b) => a + b, 0) + 1;

//     // Título
//     const title = ws.addRow([
//       `Reporte Semanal  - Mes ${monthName} Semana ${selectedWeek}`,
//     ]);
//     ws.mergeCells(1, 1, 1, totalDataCols);
//     title.font = {
//       name: "Calibri",
//       size: 16,
//       bold: true,
//       color: { argb: WHITE },
//     };
//     title.alignment = { horizontal: "center", vertical: "middle" };
//     title.fill = {
//       type: "pattern",
//       pattern: "solid",
//       fgColor: { argb: TITLE_BG },
//     };
//     ws.getRow(1).height = 30;

//     // Subtítulo
//     const subtitle = ws.addRow([
//       `Total de empleados: ${activeEmployees.length}  |  ${isActive == 0 ? "INACTIVOS" : "ACTIVOS"
//       } | P: Permisos / L: Lactancia / E: Embarazada / T: Tarde / EP: Permiso Especial `,
//     ]);
//     ws.mergeCells(2, 1, 2, totalDataCols);
//     subtitle.font = { name: "Calibri", size: 12, bold: true };
//     subtitle.alignment = { horizontal: "center", vertical: "middle" };
//     subtitle.fill = {
//       type: "pattern",
//       pattern: "solid",
//       fgColor: { argb: "F2F2F2" },
//     };
//     ws.getRow(2).height = 24;

//     // Fila 3 (grupo por día) + Fila 4 (subheaders)
//     const mainHeaderRow = [];
//     mainHeaderRow.push("", "", "", "");
//     columnsPerDay.forEach((n) => mainHeaderRow.push(...Array(n).fill("")));
//     mainHeaderRow.push("");
//     ws.addRow(mainHeaderRow);

//     const subHeaderRow = ["Item", "Código", "Empleado", "Tipo de Planilla"];
//     days.forEach((_, i) => {
//       subHeaderRow.push("(P/L/E/EP/T)", "E", "S", "H.T", "H.A");
//       for (let j = 1; j <= maxPermsByDay[i]; j++)
//         subHeaderRow.push(`P${j}S`, `P${j}E`);
//       if (includeDispatchByDay[i]) subHeaderRow.push("D");
//     });
//     subHeaderRow.push("Comentarios");
//     ws.addRow(subHeaderRow);

//     ws.mergeCells(3, 1, 4, 1);
//     ws.mergeCells(3, 2, 4, 2);
//     ws.mergeCells(3, 3, 4, 3);
//     ws.mergeCells(3, 4, 4, 4);
//     ws.mergeCells(3, totalDataCols, 4, totalDataCols);

//     const setStatic = (col, label) => {
//       const cell = ws.getCell(3, col);
//       cell.value = label;
//       cell.font = { name: "Calibri", size: 12, bold: true };
//       cell.alignment = { horizontal: "center", vertical: "middle" };
//       cell.fill = {
//         type: "pattern",
//         pattern: "solid",
//         fgColor: { argb: WHITE },
//       };
//       cell.border = {
//         top: { style: "thin" },
//         bottom: { style: "thin" },
//         left: { style: "thin" },
//         right: { style: "thin" },
//       };
//     };
//     setStatic(1, "Item");
//     setStatic(2, "Código");
//     setStatic(3, "Empleado");
//     setStatic(4, "Tipo de Planilla");
//     setStatic(totalDataCols, "Comentarios");

//     let cur = 5;
//     const dayStart = [];
//     days.forEach((d) => {
//       const span = columnsPerDay[d.idx];
//       dayStart.push(cur);
//       ws.mergeCells(3, cur, 3, cur + span - 1);
//       const c = ws.getCell(3, cur);
//       c.value = `${d.name} ${d.short}`;
//       c.font = { name: "Calibri", size: 12, bold: true };
//       c.alignment = { horizontal: "center", vertical: "middle" };
//       c.fill = {
//         type: "pattern",
//         pattern: "solid",
//         fgColor: { argb: dayHeaderColors[d.name] || WHITE },
//       };
//       c.border = {
//         top: { style: "thin" },
//         bottom: { style: "thin" },
//         left: { style: "thin" },
//         right: { style: "thin" },
//       };
//       cur += span;
//     });
//     ws.getRow(3).height = 20;

//     ws.getRow(4).font = { name: "Calibri", size: 11, bold: true };
//     ws.getRow(4).alignment = { horizontal: "center", vertical: "middle" };
//     ws.getRow(4).eachCell((cell) => {
//       cell.border = {
//         top: { style: "thin" },
//         bottom: { style: "thin" },
//         left: { style: "thin" },
//         right: { style: "thin" },
//       };
//       cell.fill = {
//         type: "pattern",
//         pattern: "solid",
//         fgColor: { argb: WHITE },
//       };
//     });

//     days.forEach((_, i) => {
//       const s = dayStart[i];
//       ws.getRow(4).getCell(s + 1).fill = {
//         type: "pattern",
//         pattern: "solid",
//         fgColor: { argb: GREEN_E },
//       };
//       ws.getRow(4).getCell(s + 2).fill = {
//         type: "pattern",
//         pattern: "solid",
//         fgColor: { argb: RED_S },
//       };
//       const m = maxPermsByDay[i];
//       for (let j = 0; j < m; j++) {
//         ws.getRow(4).getCell(s + 5 + j * 2).fill = {
//           type: "pattern",
//           pattern: "solid",
//           fgColor: { argb: GRAY_P },
//         };
//         ws.getRow(4).getCell(s + 5 + j * 2 + 1).fill = {
//           type: "pattern",
//           pattern: "solid",
//           fgColor: { argb: GRAY_P },
//         };
//       }
//       if (includeDispatchByDay[i])
//         ws.getRow(4).getCell(s + 5 + m * 2).fill = {
//           type: "pattern",
//           pattern: "solid",
//           fgColor: { argb: YELLOW_D },
//         };
//     });

//     // ====== Filas de datos ======
//     const tableRows = activeEmployees.map((emp, idx) => {
//       const row = [
//         idx + 1,
//         emp.employeeID,
//         emp.employeeName,
//         payrollMap.get(String(emp.employeeID)) || "",
//       ];
//       const weekComments = [];

//       days.forEach((d, dIdx) => {
//         const rec =
//           san.find(
//             (r) =>
//               r.employeeID === String(emp.employeeID) && r.date === d.date
//           ) || {};
//         const entry = rec.entryTime || "";
//         const exit = rec.exitTime || "";
//         const noAttendance = !entry && !exit;

//         const permList =
//           permsByEmpDate.get(`${emp.employeeID}|${d.date}`) || [];

//         const shiftCfg = getShiftConfig(emp.shiftID, d.iso);

//         // Detectar si el turno es nocturno (cruza medianoche)
//         const isNightShift = shiftCfg.endMin > 24 * 60;

//         // Minutos reales de entrada/salida (sin ajustar)
//         const entryMinRaw = timeToMinutes(entry);
//         const exitMinRaw = timeToMinutes(exit);

//         // ¿Las marcas de este día están en horario diurno?
//         // (por ejemplo entre 4:00 y 20:00, ajusta si quieres otro rango)
//         const marksInDay =
//           entryMinRaw != null &&
//           exitMinRaw != null &&
//           entryMinRaw >= 4 * 60 &&
//           exitMinRaw <= 20 * 60 &&
//           exitMinRaw > entryMinRaw;

//         // TAGS (P/L/E/EP/T)
//         const hasP = permList.length > 0;
//         const flags =
//           flagsByEmpDate.get(`${emp.employeeID}|${d.date}`) || {
//             L: 0,
//             E: 0,
//             EP: 0,
//           };
//         const hasL = flags.L === 1;
//         const hasE = flags.E === 1;
//         const hasEP = flags.EP === 1;

//         // Tarde: usamos hora de inicio distinta si es turno nocturno pero marcó de día
//         let hasT = false;
//         let startMinRef = shiftCfg.startMin;

//         // Si el turno es nocturno pero marcó de día, asumimos turno diurno genérico 07:00
//         if (isNightShift && marksInDay) {
//           startMinRef = 7 * 60; // 07:00 AM
//         }

//         if (entryMinRaw != null) {
//           hasT = entryMinRaw > startMinRef + 1;
//         }

//         const tagsParts = [];
//         if (hasP) tagsParts.push("P");
//         if (hasL) tagsParts.push("L");
//         if (hasE) tagsParts.push("E");
//         if (hasEP) tagsParts.push("EP");
//         if (hasT) tagsParts.push("T");
//         const tags = tagsParts.join(".") || "-";

//         const entryF = fmt12smart(entry, { bias: "entry" });
//         const exitF = fmt12smart(exit, { bias: "exit", refEntry: entry });

//         // ===== Cálculo H.T / H.A =====
//         let ht = "";
//         let ha = "";

//         // Configuración de turno a usar
//         let cfgToUse = shiftCfg;

//         // Si el turno es nocturno pero las marcas son claramente diurnas,
//         // usamos una jornada genérica 07:00–16:45
//         if (isNightShift && marksInDay) {
//           const startMin = 7 * 60;           // 07:00
//           const endMin = 16 * 60 + 45;       // 16:45
//           const lunchHours = LUNCH_HOURS_DEFAULT; // 0.75
//           const hrsLaborables = (endMin - startMin) / 60 - lunchHours;
//           const lunchThresholdMin = startMin + LUNCH_OFFSET_MIN;

//           cfgToUse = { startMin, endMin, lunchHours, lunchThresholdMin, hrsLaborables };
//         }

//         // Reglas de cálculo
//         const hasEntry = !!entry;
//         const hasExit = !!exit;
//         const hasPerm = permList.length > 0;

//         if (!hasEntry && !hasExit && !hasPerm) {
//           // Sin marcas ni permisos → no mostramos nada
//           ht = "";
//           ha = "";
//         } else if ((hasEntry && !hasExit) || (!hasEntry && hasExit)) {
//           // Solo entrada o solo salida → 0/0 sin ausencia
//           ht = 0;
//           ha = 0;
//         } else {
//           // Caso normal: usamos computeDaySummary con cfgToUse
//           const exceptionTime =
//             exceptionTimeMap.get(String(emp.employeeID)) || 0;
//           const { hrsTrabajadas, hrsAusencia } = computeDaySummary(
//             entry,
//             exit,
//             permList,
//             exceptionTime,
//             cfgToUse
//           );

//           ht =
//             hrsTrabajadas === null || hrsTrabajadas === undefined
//               ? ""
//               : Number(hrsTrabajadas.toFixed(2));
//           ha =
//             hrsAusencia === null || hrsAusencia === undefined
//               ? ""
//               : Number(hrsAusencia.toFixed(2));
//         }

//         // Agregamos a la fila
//         row.push(tags);
//         row.push(entryF);
//         row.push(exitF);
//         row.push(ht === "" ? "" : ht);
//         row.push(ha === "" ? "" : ha);


//         // ====== Permisos (celdas PjS / PjE) ======
//         const list = permList.slice(0, maxPermsByDay[dIdx]);
//         const firstSolicIdx = list.findIndex(
//           (p) => p.request === 1 && !p.exitFmt && !p.entryFmt
//         );

//         for (let iPerm = 0; iPerm < maxPermsByDay[dIdx]; iPerm++) {
//           const p = list[iPerm];
//           if (!p) {
//             row.push("-", "-");
//             continue;
//           }

//           if (noAttendance && iPerm === firstSolicIdx) {
//             row.push(NO_SHOW, NO_SHOW);
//             weekComments.push(
//               `${d.abbr} P${iPerm + 1}: No vino, no marco asistencia`
//             );
//             continue;
//           }

//           if (p.request === 0 && !p.exitFmt && !p.entryFmt) {
//             const reason = p.comment ? ` – Justificación: ${p.comment}` : "";
//             row.push(`${NO_HOUR}${reason}`, `${NO_HOUR}${reason}`);
//             weekComments.push(
//               `${d.abbr} P${iPerm + 1}: sin hora${p.comment ? ` (Justificación: ${p.comment})` : ""
//               }`
//             );
//             continue;
//           }

//           const sVal = fmt12smart(p.exitFmt, { bias: "auto" });
//           const eVal = fmt12smart(p.entryFmt, { bias: "auto" });
//           row.push(sVal, eVal);
//         }

//         if (includeDispatchByDay[dIdx]) {
//           row.push(
//             fmt12smart(rec.dispatchingTime || "", { bias: "auto" })
//           );
//         }

//         if (rec.exitComment)
//           weekComments.push(`${d.abbr} Salida: ${rec.exitComment}`);
//         if (rec.dispatchingComment)
//           weekComments.push(`${d.abbr} Despacho: ${rec.dispatchingComment}`);
//       });

//       row.push(weekComments.join(" | ") || "");
//       return row;
//     });

//     tableRows.forEach((r) => ws.addRow(r));

//     ws.eachRow((row, rNum) => {
//       if (rNum > 4) {
//         row.eachCell((cell) => {
//           cell.font = { name: "Calibri", size: 11 };
//           cell.alignment = { horizontal: "center", vertical: "middle" };
//           cell.border = {
//             top: { style: "thin" },
//             bottom: { style: "thin" },
//             left: { style: "thin" },
//             right: { style: "thin" },
//           };
//         });
//         row.fill = {
//           type: "pattern",
//           pattern: "solid",
//           fgColor: { argb: rNum % 2 === 1 ? "F5F5F5" : "FFFFFF" },
//         };
//         row.height = 20;
//       }
//     });

//     const widths = [
//       { width: 10 },
//       { width: 10 },
//       { width: 30 },
//       { width: 16 },
//     ];
//     days.forEach((_, i) => {
//       widths.push({ width: 12 });
//       widths.push({ width: 12 });
//       widths.push({ width: 12 });
//       widths.push({ width: 10 });
//       widths.push({ width: 10 });
//       for (let j = 1; j <= maxPermsByDay[i]; j++) {
//         widths.push({ width: 12 });
//         widths.push({ width: 12 });
//       }
//       if (includeDispatchByDay[i]) widths.push({ width: 12 });
//     });
//     widths.push({ width: 60 });
//     ws.columns = widths;

//     const buf = await workbook.xlsx.writeBuffer();
//     const filename = `asistencia_semanal_${isActive == 0 ? "inactivos" : "activos"
//       }_semana${selectedWeek}_${year}.xlsx`;
//     res.setHeader(
//       "Content-Disposition",
//       `attachment; filename="${filename}"`
//     );
//     res.setHeader(
//       "Content-Type",
//       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
//     );
//     res.send(buf);
//   } catch (err) {
//     console.error(
//       "Error al exportar asistencia semanal (H.T/H.A + permisos/Despacho):",
//       err.stack || err.message
//     );
//     res.status(500).send({
//       message: `Error interno del servidor al generar el archivo Excel: ${err.message}`,
//     });
//   }
// };

// exports.exportWeeklyAttendance = async (req, res) => {
//   try {
//     const { weeklyAttendance, selectedMonth, selectedWeek, isActive } = req.body;

//     //  Permitimos que weeklyAttendance venga vacío, porque ahora completamos desde BD
//     if (!weeklyAttendance || !Array.isArray(weeklyAttendance))
//       throw new Error("weeklyAttendance is missing or invalid.");
//     if (!selectedMonth || typeof selectedMonth !== "string")
//       throw new Error("selectedMonth is missing or invalid.");
//     if (!selectedWeek || typeof selectedWeek !== "string")
//       throw new Error("selectedWeek is missing or invalid.");

//     // ====== Helpers de tiempo ======
//     const parseTime = (t) =>
//       dayjs(t, ["HH:mm:ss", "H:mm:ss", "hh:mm:ss A", "h:mm:ss A"], true);

//     const fmt12smart = (t, { bias = "auto", refEntry } = {}) => {
//       if (!t) return "-";
//       let d = dayjs(
//         t,
//         [
//           "HH:mm:ss",
//           "H:mm:ss",
//           "hh:mm:ss A",
//           "h:mm:ss A",
//           "YYYY-MM-DD HH:mm:ss",
//           "YYYY-MM-DD hh:mm:ss A",
//         ],
//         true
//       );
//       if (!d.isValid()) return String(t);

//       const raw = String(t);
//       const hasMeridiem = /AM|PM/i.test(raw);
//       if (!hasMeridiem && bias === "exit" && refEntry) {
//         const e = dayjs(
//           refEntry,
//           ["HH:mm:ss", "H:mm:ss", "hh:mm:ss A", "h:mm:ss A"],
//           true
//         );
//         if (e.isValid()) {
//           const hrEntry = e.hour();
//           const hrExit = d.hour();
//           if (hrEntry >= 5 && hrEntry <= 9 && hrExit >= 1 && hrExit <= 6)
//             d = d.add(12, "hour");
//         }
//       }
//       return d.format("hh:mm:ss A");
//     };

//     const timeToMinutes = (t) => {
//       if (!t) return null;
//       const d = parseTime(t);
//       if (!d.isValid()) return null;
//       return d.hour() * 60 + d.minute();
//     };

//     const mergeIntervals = (intervals) => {
//       if (!intervals.length) return [];
//       const sorted = intervals
//         .slice()
//         .sort((a, b) => a[0] - b[0] || a[1] - b[1]);
//       const result = [];
//       let [curStart, curEnd] = sorted[0];
//       for (let i = 1; i < sorted.length; i++) {
//         const [s, e] = sorted[i];
//         if (s <= curEnd) {
//           if (e > curEnd) curEnd = e;
//         } else {
//           result.push([curStart, curEnd]);
//           curStart = s;
//           curEnd = e;
//         }
//       }
//       result.push([curStart, curEnd]);
//       return result;
//     };

//     const totalMinutes = (intervals) => {
//       const merged = mergeIntervals(intervals);
//       return merged.reduce((sum, [s, e]) => sum + (e - s), 0);
//     };

//     const coverageInSegment = (intervals, segStart, segEnd) => {
//       if (segStart == null || segEnd == null || segEnd <= segStart) return 0;
//       const clipped = [];
//       for (const [s, e] of intervals) {
//         const cs = Math.max(segStart, s);
//         const ce = Math.min(segEnd, e);
//         if (ce > cs) clipped.push([cs, ce]);
//       }
//       return totalMinutes(clipped);
//     };

//     // ====== Constantes / colores ======
//     const TITLE_BG = "1F3864";
//     const WHITE = "FFFFFF";
//     const GREEN_E = "C6EFCE";
//     const RED_S = "FFC7CE";
//     const GRAY_P = "D6DCE4";
//     const YELLOW_D = "FFEB9C";
//     const NO_HOUR = "(sin hora)";
//     const NO_SHOW = "(No vino, no marco asistencia)";

//     // Almuerzo genérico
//     const LUNCH_HOURS_DEFAULT = 0.75; // 45 minutos
//     const LUNCH_OFFSET_MIN = 5.75 * 60; // 5h45 después de la entrada

//     const dayHeaderColors = {
//       Lunes: "FFFFFF",
//       Martes: "FFC0CB",
//       Miércoles: "FFFFFF",
//       Jueves: "FFFFFF",
//       Viernes: "FFFFFF",
//       Sábado: "D3D3D3",
//       Domingo: "D3D3D3",
//     };

//     // ====== Semana ISO (ANTES de buscar empleados) ======
//     let startOfWeek;
//     if (weeklyAttendance.length > 0) {
//       const firstDate = weeklyAttendance.map((r) => String(r.date)).sort()[0];
//       startOfWeek = dayjs(firstDate).startOf("isoWeek");
//     } else {
//       const yearNow = dayjs().year();
//       startOfWeek = dayjs()
//         .year(yearNow)
//         .isoWeek(parseInt(selectedWeek))
//         .startOf("isoWeek");
//     }
//     const endOfWeek = startOfWeek.add(6, "day");

//     const days = Array.from({ length: 7 }, (_, i) => {
//       const d = startOfWeek.add(i, "day");
//       return {
//         idx: i,
//         iso: d.isoWeekday(), // 1 = lunes ... 7 = domingo
//         date: d.format("YYYY-MM-DD"),
//         name:
//           d.format("dddd").charAt(0).toUpperCase() + d.format("dddd").slice(1),
//         short: d.format("DD/MM"),
//         abbr: d.format("ddd"),
//       };
//     });

//     // ====== Empleados según isActive + despedidos dentro de la semana ======
//     const [activeEmployees] = await db.query(
//       `
//       SELECT 
//         e.employeeID,
//         e.shiftID,
//         CONCAT(e.firstName,' ',COALESCE(e.middleName,''),' ',e.lastName) AS employeeName,
//         CASE WHEN d.employeeID IS NOT NULL THEN 1 ELSE 0 END AS isDismissedWeek,
//         d.dateDismissal
//       FROM employees_emp e
//       LEFT JOIN (
//         SELECT hd.employeeID, MAX(DATE(hd.dateDismissal)) AS dateDismissal
//         FROM h_dismissal_emp hd
//         WHERE DATE(hd.dateDismissal) BETWEEN ? AND ?
//         GROUP BY hd.employeeID
//       ) d ON d.employeeID = e.employeeID
//       WHERE e.isActive = ?
//          OR d.employeeID IS NOT NULL
//       ORDER BY e.employeeID
//       `,
//       [startOfWeek.format("YYYY-MM-DD"), endOfWeek.format("YYYY-MM-DD"), Number(isActive)]
//     );

//     const activeSet = new Set(activeEmployees.map((e) => String(e.employeeID)));

//     // =====================================================================
//     // ✅ CLAVE: NO depender del frontend
//     // Traemos marcajes reales desde BD para TODOS los empleados del reporte
//     // =====================================================================
//     const ids = [...activeSet];
//     let weeklyAttendanceFinal = Array.isArray(weeklyAttendance) ? weeklyAttendance.slice() : [];

//     if (ids.length > 0) {
//       // 1) asistencia real (h_attendance_emp)
//       const [attRows] = await db.query(
//         `
//         SELECT 
//           a.employeeID,
//           DATE(a.date) AS date,
//           TIME_FORMAT(a.entryTime,'%H:%i:%s') AS entryTime,
//           TIME_FORMAT(a.exitTime, '%H:%i:%s') AS exitTime,
//           COALESCE(a.comment,'') AS exitComment
//         FROM h_attendance_emp a
//         WHERE DATE(a.date) BETWEEN ? AND ?
//           AND a.employeeID IN (${ids.map(() => "?").join(",")})
//         `,
//         [startOfWeek.format("YYYY-MM-DD"), endOfWeek.format("YYYY-MM-DD"), ...ids]
//       );

//       // 2) despacho real (dispatching_emp)
//       const [dispRows] = await db.query(
//         `
//         SELECT
//           d.employeeID,
//           DATE(d.date) AS date,
//           TIME_FORMAT(d.exitTimeComplete,'%H:%i:%s') AS dispatchingTime,
//           COALESCE(d.comment,'') AS dispatchingComment
//         FROM dispatching_emp d
//         WHERE DATE(d.date) BETWEEN ? AND ?
//           AND d.employeeID IN (${ids.map(() => "?").join(",")})
//         `,
//         [startOfWeek.format("YYYY-MM-DD"), endOfWeek.format("YYYY-MM-DD"), ...ids]
//       );

//       const keyOf = (empId, dateStr) => `${String(empId)}|${dayjs(String(dateStr)).format("YYYY-MM-DD")}`;

//       const merged = new Map();

//       // A) cargar lo que venga del frontend
//       for (const r of weeklyAttendanceFinal) {
//         const key = keyOf(r.employeeID, r.date);
//         merged.set(key, {
//           employeeID: String(r.employeeID),
//           date: dayjs(String(r.date)).format("YYYY-MM-DD"),
//           entryTime: String(r.entryTime ?? ""),
//           exitTime: String(r.exitTime ?? ""),
//           dispatchingTime: String(r.dispatchingTime ?? ""),
//           dispatchingComment: String(r.dispatchingComment ?? ""),
//           exitComment: String(r.exitComment ?? ""),
//         });
//       }

//       // B) completar con asistencia de BD
//       for (const r of attRows || []) {
//         const key = keyOf(r.employeeID, r.date);
//         const prev = merged.get(key) || {
//           employeeID: String(r.employeeID),
//           date: dayjs(String(r.date)).format("YYYY-MM-DD"),
//           entryTime: "",
//           exitTime: "",
//           dispatchingTime: "",
//           dispatchingComment: "",
//           exitComment: "",
//         };

//         merged.set(key, {
//           ...prev,
//           entryTime: prev.entryTime || (r.entryTime || ""),
//           exitTime: prev.exitTime || (r.exitTime || ""),
//           exitComment: prev.exitComment || (r.exitComment || ""),
//         });
//       }

//       // C) completar con despacho de BD
//       for (const r of dispRows || []) {
//         const key = keyOf(r.employeeID, r.date);
//         const prev = merged.get(key) || {
//           employeeID: String(r.employeeID),
//           date: dayjs(String(r.date)).format("YYYY-MM-DD"),
//           entryTime: "",
//           exitTime: "",
//           dispatchingTime: "",
//           dispatchingComment: "",
//           exitComment: "",
//         };

//         merged.set(key, {
//           ...prev,
//           dispatchingTime: prev.dispatchingTime || (r.dispatchingTime || ""),
//           dispatchingComment: prev.dispatchingComment || (r.dispatchingComment || ""),
//         });
//       }

//       weeklyAttendanceFinal = Array.from(merged.values());
//     }

//     // ====== Mapa de Tipo de Planilla ======
//     const [payrollRows] = await db.query(`
//       SELECT e.employeeID, COALESCE(pt.payrollName,'') AS payrollName
//       FROM employees_emp e
//       LEFT JOIN payrolltype_emp pt ON pt.payrollTypeID = e.payrollTypeID
//     `);
//     const payrollMap = new Map(
//       payrollRows.map((r) => [String(r.employeeID), r.payrollName || ""])
//     );

//     // ====== Detalle de turnos por día (detailshift_emp) ======
//     const [shiftDetailRows] = await db.query(`
//       SELECT shiftID, day, startTime, endTime
//       FROM detailsshift_emp
//     `);

//     const dayNameToIso = {
//       Monday: 1,
//       Tuesday: 2,
//       Wednesday: 3,
//       Thursday: 4,
//       Friday: 5,
//       Saturday: 6,
//       Sunday: 7,
//     };

//     const shiftDetailMap = new Map();
//     for (const r of shiftDetailRows) {
//       const isoDay = dayNameToIso[r.day];
//       if (!isoDay) continue;
//       let s = timeToMinutes(r.startTime);
//       let e = timeToMinutes(r.endTime);
//       if (s == null || e == null) continue;
//       if (e <= s) e += 24 * 60;
//       shiftDetailMap.set(`${r.shiftID}|${isoDay}`, { startMin: s, endMin: e });
//     }

//     const getShiftConfig = (shiftId, isoDay) => {
//       const key = `${shiftId}|${isoDay}`;
//       const base = shiftDetailMap.get(key);

//       if (base) {
//         const { startMin, endMin } = base;
//         const lunchHours = LUNCH_HOURS_DEFAULT;
//         const hrsLaborables = (endMin - startMin) / 60 - lunchHours;
//         const lunchThresholdMin = startMin + LUNCH_OFFSET_MIN;
//         return { startMin, endMin, lunchHours, lunchThresholdMin, hrsLaborables };
//       }

//       const startMin = 7 * 60;
//       const endMin = 16 * 60 + 45;
//       const lunchHours = LUNCH_HOURS_DEFAULT;
//       const hrsLaborables = (endMin - startMin) / 60 - lunchHours;
//       const lunchThresholdMin = startMin + LUNCH_OFFSET_MIN;
//       return { startMin, endMin, lunchHours, lunchThresholdMin, hrsLaborables };
//     };

//     // ====== Mapa de exceptionTime ======
//     const [excRows] = await db.query(`
//       SELECT e.employeeID, COALESCE(ex.exceptionTime, 0) AS exceptionTime
//       FROM employees_emp e
//       LEFT JOIN exceptions_emp ex ON ex.exceptionID = e.exceptionID
//       WHERE e.employeeID IN (${[...activeSet].map((id) => db.escape(id)).join(",")})
//     `);
//     const exceptionTimeMap = new Map(
//       excRows.map((r) => [String(r.employeeID), Number(r.exceptionTime || 0)])
//     );

//     // ====== Indexar asistencias (YA con weeklyAttendanceFinal) ======
//     const attMap = new Map(
//       weeklyAttendanceFinal
//         .filter((r) => activeSet.has(String(r.employeeID)))
//         .map((r) => {
//           const key = `${String(r.employeeID)}|${String(r.date)}`;
//           return [
//             key,
//             {
//               entryTime: String(r.entryTime ?? ""),
//               exitTime: String(r.exitTime ?? ""),
//               dispatchingTime: String(r.dispatchingTime ?? ""),
//               dispatchingComment: String(r.dispatchingComment ?? ""),
//               exitComment: String(r.exitComment ?? ""),
//             },
//           ];
//         })
//     );

//     // ====== Grid base: todos los empleados × 7 días ======
//     const san = [];
//     for (const emp of activeEmployees) {
//       const empId = String(emp.employeeID);
//       for (const d of days) {
//         const key = `${empId}|${d.date}`;
//         const rec =
//           attMap.get(key) || {
//             entryTime: "",
//             exitTime: "",
//             dispatchingTime: "",
//             dispatchingComment: "",
//             exitComment: "",
//           };
//         san.push({
//           employeeID: empId,
//           employeeName: String(emp.employeeName ?? ""),
//           date: d.date,
//           ...rec,
//         });
//       }
//     }

//     // ====== Permisos (aprobados) de la semana ======
//     const [permWeekRows] = await db.query(
//       `
//       SELECT employeeID,
//              DATE(date) AS date,
//              TIME_FORMAT(exitPermission,  '%H:%i:%s') AS exitFmt,
//              TIME_FORMAT(entryPermission, '%H:%i:%s') AS entryFmt,
//              request,
//              comment,
//              lunchTime,
//              createdDate,
//              permissionID
//       FROM permissionattendance_emp
//       WHERE date BETWEEN ? AND ?
//         AND isApproved = 1
//       ORDER BY employeeID, date, createdDate, permissionID
//       `,
//       [startOfWeek.format("YYYY-MM-DD"), endOfWeek.format("YYYY-MM-DD")]
//     );

//     const permsByEmpDate = new Map();
//     for (const p of permWeekRows) {
//       const emp = String(p.employeeID);
//       if (!activeSet.has(emp)) continue;
//       const key = `${emp}|${dayjs(p.date).format("YYYY-MM-DD")}`;
//       if (!permsByEmpDate.has(key)) permsByEmpDate.set(key, []);
//       permsByEmpDate.get(key).push({
//         request: p.request == 0 || p.request === "0" ? 0 : 1,
//         exitFmt: p.exitFmt || "",
//         entryFmt: p.entryFmt || "",
//         comment: p.comment || "",
//         lunchTime: p.lunchTime,
//       });
//     }

//     // ====== Excepciones por día (L, E, EP) ======
//     let hasBridge = false;
//     let bridgeRows = [];
//     try {
//       const [rowsBridge] = await db.query(`
//         SELECT 
//           ee.employeeID,
//           ee.isActive,
//           ee.startDate,
//           ee.endDate,
//           UPPER(ex.exceptionName) AS exceptionName
//         FROM employee_exceptions ee
//         JOIN exceptions_emp ex ON ex.exceptionID = ee.exceptionID
//         WHERE ee.employeeID IN (${[...activeSet].map((id) => db.escape(id)).join(",")})
//       `);
//       bridgeRows = rowsBridge || [];
//       hasBridge = bridgeRows.length > 0;
//     } catch (_) {
//       hasBridge = false;
//       bridgeRows = [];
//     }

//     const flagsByEmpDate = new Map();
//     if (hasBridge) {
//       for (const empId of activeSet) {
//         const myExc = bridgeRows.filter((r) => String(r.employeeID) === String(empId));
//         for (const d of days) {
//           const applies = (rec) => {
//             const active =
//               rec.isActive == null || rec.isActive === 1 || rec.isActive === "1";
//             const sdOk =
//               !rec.startDate || dayjs(d.date).isSameOrAfter(dayjs(rec.startDate), "day");
//             const edOk =
//               !rec.endDate || dayjs(d.date).isSameOrBefore(dayjs(rec.endDate), "day");
//             return active && sdOk && edOk;
//           };
//           const norm = (x) => (x || "").toString().trim().toUpperCase();
//           const L = myExc.some((r) => applies(r) && norm(r.exceptionName) === "LACTANCIA") ? 1 : 0;
//           const E = myExc.some((r) => applies(r) && norm(r.exceptionName) === "EMBARAZADA") ? 1 : 0;
//           const EP = myExc.some((r) => applies(r) && norm(r.exceptionName) === "PERMISO ESPECIAL") ? 1 : 0;
//           flagsByEmpDate.set(`${empId}|${d.date}`, { L, E, EP });
//         }
//       }
//     } else {
//       const [fkRows] = await db.query(`
//         SELECT 
//           e.employeeID,
//           UPPER(COALESCE(ex.exceptionName,'')) AS exName
//         FROM employees_emp e
//         LEFT JOIN exceptions_emp ex ON ex.exceptionID = e.exceptionID
//         WHERE e.employeeID IN (${[...activeSet].map((id) => db.escape(id)).join(",")})
//       `);
//       const fkMap = new Map(fkRows.map((r) => [String(r.employeeID), r.exName || ""]));
//       for (const empId of activeSet) {
//         const exName = (fkMap.get(String(empId)) || "").toUpperCase();
//         const L = exName === "LACTANCIA" ? 1 : 0;
//         const E = exName === "EMBARAZADA" ? 1 : 0;
//         const EP = 0;
//         for (const d of days) flagsByEmpDate.set(`${empId}|${d.date}`, { L, E, EP });
//       }
//     }

//     // ====== Función para cálculo diario (H.T / H.A) ======
//     const computeDaySummary = (entry, exit, permsList, exceptionTime = 0, shiftCfg) => {
//       const { startMin, endMin, lunchHours, lunchThresholdMin, hrsLaborables } = shiftCfg;
//       const crossesMidnight = endMin > 24 * 60;

//       const adjustForShift = (min) => {
//         if (min == null) return null;
//         if (crossesMidnight && min < startMin && min <= 12 * 60) min += 24 * 60;
//         return min;
//       };

//       let entryMin = entry ? adjustForShift(timeToMinutes(entry)) : null;
//       let exitMin = exit ? adjustForShift(timeToMinutes(exit)) : null;

//       const entryMinClamped =
//         entryMin == null ? null : Math.min(Math.max(entryMin, startMin), endMin);
//       const exitMinClamped =
//         exitMin == null ? null : Math.min(Math.max(exitMin, startMin), endMin);

//       const permisoIntervalsRaw = [];
//       let permisoLunchFlag = false;

//       for (const p of permsList) {
//         let s = adjustForShift(timeToMinutes(p.exitFmt));
//         let e = adjustForShift(timeToMinutes(p.entryFmt));
//         if (s == null || e == null) continue;
//         if (e <= s) continue;
//         permisoIntervalsRaw.push([s, e]);
//         if (Number(p.lunchTime || 0) === 1) permisoLunchFlag = true;
//       }

//       const permisoMinInShift = coverageInSegment(permisoIntervalsRaw, startMin, endMin);
//       let hrsPermiso = permisoMinInShift / 60.0;
//       if (permisoLunchFlag && hrsPermiso > 0) hrsPermiso = Math.max(hrsPermiso - lunchHours, 0);

//       const anyLunchFlag = permisoLunchFlag;
//       const soloEntradaSinSalida = !!entry && !exit;
//       const soloSalidaSinEntrada = !entry && !!exit;
//       const caso3 = soloEntradaSinSalida || soloSalidaSinEntrada;

//       let hrsAusente_entrada = 0;
//       let hrsAusente_salida = 0;
//       let hrsAusencia = 0;
//       let hrsTrabajadas = 0;

//       if (caso3) {
//         hrsTrabajadas = 0;
//         hrsAusencia = 0;
//       } else {
//         const noEntry = !entry;
//         const noExit = !exit;

//         if (noEntry && noExit && permisoMinInShift === 0) {
//           const netPermiso = hrsPermiso;
//           hrsAusencia = Math.max(hrsLaborables - netPermiso, 0);
//         } else {
//           if (entryMinClamped != null && entryMinClamped > startMin) {
//             const totalGapMin = entryMinClamped - startMin;
//             const coveredMin = coverageInSegment(permisoIntervalsRaw, startMin, entryMinClamped);
//             let baseMin = totalGapMin - coveredMin;
//             if (baseMin < 0) baseMin = 0;

//             if (baseMin > 0 && entryMinClamped >= lunchThresholdMin && !anyLunchFlag) {
//               baseMin -= lunchHours * 60;
//               if (baseMin < 0) baseMin = 0;
//             }

//             hrsAusente_entrada = baseMin / 60.0;
//           }

//           if (exitMinClamped != null && exitMinClamped < endMin) {
//             const totalGapMin = endMin - exitMinClamped;
//             const coveredAfterExit = coverageInSegment(permisoIntervalsRaw, exitMinClamped, endMin);
//             let effectiveGapMin = totalGapMin - coveredAfterExit;
//             if (effectiveGapMin < 0) effectiveGapMin = 0;

//             let baseHours = effectiveGapMin / 60.0 - Number(exceptionTime || 0);
//             if (baseHours < 0) baseHours = 0;

//             hrsAusente_salida = baseHours;
//           }

//           hrsAusencia = hrsAusente_entrada + hrsAusente_salida;
//         }

//         hrsTrabajadas = Math.max(hrsLaborables - (hrsPermiso + hrsAusencia), 0);
//       }

//       return { hrsTrabajadas, hrsAusencia };
//     };

//     // ====== Máximo de permisos por día ======
//     const maxPermsByDay = Array(7).fill(0);
//     for (const [key, list] of permsByEmpDate.entries()) {
//       const [, dateStr] = key.split("|");
//       const idx = dayjs(dateStr).diff(startOfWeek, "day");
//       if (idx < 0 || idx > 6) continue;
//       maxPermsByDay[idx] = Math.max(maxPermsByDay[idx], Math.min(list.length, 5));
//     }

//     // ====== ¿Hay despacho por día? ======
//     const includeDispatchByDay = Array(7).fill(false);
//     for (let i = 0; i < 7; i++) {
//       includeDispatchByDay[i] = san.some(
//         (r) => r.date === days[i].date && r.dispatchingTime && activeSet.has(r.employeeID)
//       );
//     }

//     // ====== Si no hay empleados ======
//     if (activeEmployees.length === 0) {
//       const workbook0 = new ExcelJS.Workbook();
//       const w = workbook0.addWorksheet("Asistencia Semanal");
//       w.addRow([`Reporte Semanal - ${isActive == 0 ? "INACTIVOS" : "ACTIVOS"} - (sin empleados)`]);
//       w.mergeCells(1, 1, 1, 5);
//       w.getRow(1).font = { bold: true };
//       w.addRow(["Item", "Código", "Empleado", "Tipo de Planilla", "Comentarios"]);
//       w.columns = [{ width: 10 }, { width: 10 }, { width: 30 }, { width: 18 }, { width: 60 }];
//       const buf0 = await workbook0.xlsx.writeBuffer();
//       const filename0 = `asistencia_semanal_${isActive == 0 ? "inactivos" : "activos"}_semana${selectedWeek}_${dayjs().year()}.xlsx`;
//       res.setHeader("Content-Disposition", `attachment; filename="${filename0}"`);
//       res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
//       return res.send(buf0);
//     }

//     // ====== Excel (estructura con H.T / H.A + permisos + despacho) ======
//     const workbook = new ExcelJS.Workbook();
//     const ws = workbook.addWorksheet("Asistencia Semanal");

//     const year = dayjs().year();
//     const monthName = dayjs().month(parseInt(selectedMonth)).format("MMMM").toUpperCase();

//     const columnsPerDay = days.map(
//       (_, i) => 1 + 4 + maxPermsByDay[i] * 2 + (includeDispatchByDay[i] ? 1 : 0)
//     );
//     const totalDataCols = 4 + columnsPerDay.reduce((a, b) => a + b, 0) + 1;

//     const title = ws.addRow([`Reporte Semanal  - Mes ${monthName} Semana ${selectedWeek}`]);
//     ws.mergeCells(1, 1, 1, totalDataCols);
//     title.font = { name: "Calibri", size: 16, bold: true, color: { argb: WHITE } };
//     title.alignment = { horizontal: "center", vertical: "middle" };
//     title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: TITLE_BG } };
//     ws.getRow(1).height = 30;

//     const subtitle = ws.addRow([
//       `Total de empleados: ${activeEmployees.length}  |  ${isActive == 0 ? "INACTIVOS" : "ACTIVOS"} | P: Permisos / L: Lactancia / E: Embarazada / T: Tarde / EP: Permiso Especial / DP: Despido`,
//     ]);
//     ws.mergeCells(2, 1, 2, totalDataCols);
//     subtitle.font = { name: "Calibri", size: 12, bold: true };
//     subtitle.alignment = { horizontal: "center", vertical: "middle" };
//     subtitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "F2F2F2" } };
//     ws.getRow(2).height = 24;

//     const mainHeaderRow = [];
//     mainHeaderRow.push("", "", "", "");
//     columnsPerDay.forEach((n) => mainHeaderRow.push(...Array(n).fill("")));
//     mainHeaderRow.push("");
//     ws.addRow(mainHeaderRow);

//     const subHeaderRow = ["Item", "Código", "Empleado", "Tipo de Planilla"];
//     days.forEach((_, i) => {
//       subHeaderRow.push("(P/L/E/EP/T/DP)", "E", "S", "H.T", "H.A");
//       for (let j = 1; j <= maxPermsByDay[i]; j++) subHeaderRow.push(`P${j}S`, `P${j}E`);
//       if (includeDispatchByDay[i]) subHeaderRow.push("D");
//     });
//     subHeaderRow.push("Comentarios");
//     ws.addRow(subHeaderRow);

//     ws.mergeCells(3, 1, 4, 1);
//     ws.mergeCells(3, 2, 4, 2);
//     ws.mergeCells(3, 3, 4, 3);
//     ws.mergeCells(3, 4, 4, 4);
//     ws.mergeCells(3, totalDataCols, 4, totalDataCols);

//     const setStatic = (col, label) => {
//       const cell = ws.getCell(3, col);
//       cell.value = label;
//       cell.font = { name: "Calibri", size: 12, bold: true };
//       cell.alignment = { horizontal: "center", vertical: "middle" };
//       cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: WHITE } };
//       cell.border = {
//         top: { style: "thin" },
//         bottom: { style: "thin" },
//         left: { style: "thin" },
//         right: { style: "thin" },
//       };
//     };
//     setStatic(1, "Item");
//     setStatic(2, "Código");
//     setStatic(3, "Empleado");
//     setStatic(4, "Tipo de Planilla");
//     setStatic(totalDataCols, "Comentarios");

//     let cur = 5;
//     const dayStart = [];
//     days.forEach((d) => {
//       const span = columnsPerDay[d.idx];
//       dayStart.push(cur);
//       ws.mergeCells(3, cur, 3, cur + span - 1);
//       const c = ws.getCell(3, cur);
//       c.value = `${d.name} ${d.short}`;
//       c.font = { name: "Calibri", size: 12, bold: true };
//       c.alignment = { horizontal: "center", vertical: "middle" };
//       c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: dayHeaderColors[d.name] || WHITE } };
//       c.border = {
//         top: { style: "thin" },
//         bottom: { style: "thin" },
//         left: { style: "thin" },
//         right: { style: "thin" },
//       };
//       cur += span;
//     });
//     ws.getRow(3).height = 20;

//     ws.getRow(4).font = { name: "Calibri", size: 11, bold: true };
//     ws.getRow(4).alignment = { horizontal: "center", vertical: "middle" };
//     ws.getRow(4).eachCell((cell) => {
//       cell.border = {
//         top: { style: "thin" },
//         bottom: { style: "thin" },
//         left: { style: "thin" },
//         right: { style: "thin" },
//       };
//       cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: WHITE } };
//     });

//     days.forEach((_, i) => {
//       const s = dayStart[i];
//       ws.getRow(4).getCell(s + 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: GREEN_E } };
//       ws.getRow(4).getCell(s + 2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: RED_S } };
//       const m = maxPermsByDay[i];
//       for (let j = 0; j < m; j++) {
//         ws.getRow(4).getCell(s + 5 + j * 2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRAY_P } };
//         ws.getRow(4).getCell(s + 5 + j * 2 + 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRAY_P } };
//       }
//       if (includeDispatchByDay[i]) {
//         ws.getRow(4).getCell(s + 5 + m * 2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: YELLOW_D } };
//       }
//     });

//     // ====== Filas de datos ======
//     const tableRows = activeEmployees.map((emp, idx) => {
//       const row = [idx + 1, emp.employeeID, emp.employeeName, payrollMap.get(String(emp.employeeID)) || ""];
//       const weekComments = [];

//       days.forEach((d, dIdx) => {
//         const rec =
//           san.find((r) => r.employeeID === String(emp.employeeID) && r.date === d.date) || {};

//         const dismissedDate = emp.dateDismissal ? dayjs(emp.dateDismissal).format("YYYY-MM-DD") : null;

//         // ✅ Regla final:
//         // - Antes del despido: normal
//         // - Día del despido: normal + DP
//         // - Después del despido: limpiar horas/permisos/despacho, pero DP se mantiene
//         const isAfterDismissalDay =
//           dismissedDate && dayjs(d.date).isAfter(dayjs(dismissedDate), "day");

//         const isDismissalOrAfter =
//           dismissedDate &&
//           (dayjs(d.date).isSame(dayjs(dismissedDate), "day") ||
//             dayjs(d.date).isAfter(dayjs(dismissedDate), "day"));

//         const entry = isAfterDismissalDay ? "" : (rec.entryTime || "");
//         const exit = isAfterDismissalDay ? "" : (rec.exitTime || "");
//         const noAttendance = !entry && !exit;

//         const permList = isAfterDismissalDay
//           ? []
//           : (permsByEmpDate.get(`${emp.employeeID}|${d.date}`) || []);

//         const shiftCfg = getShiftConfig(emp.shiftID, d.iso);

//         const isNightShift = shiftCfg.endMin > 24 * 60;
//         const entryMinRaw = timeToMinutes(entry);
//         const exitMinRaw = timeToMinutes(exit);

//         const marksInDay =
//           entryMinRaw != null &&
//           exitMinRaw != null &&
//           entryMinRaw >= 4 * 60 &&
//           exitMinRaw <= 20 * 60 &&
//           exitMinRaw > entryMinRaw;

//         // TAGS
//         const hasP = permList.length > 0;
//         const flags = flagsByEmpDate.get(`${emp.employeeID}|${d.date}`) || { L: 0, E: 0, EP: 0 };
//         const hasL = flags.L === 1;
//         const hasE = flags.E === 1;
//         const hasEP = flags.EP === 1;

//         let hasT = false;
//         let startMinRef = shiftCfg.startMin;

//         if (isNightShift && marksInDay) startMinRef = 7 * 60;
//         if (entryMinRaw != null) hasT = entryMinRaw > startMinRef + 1;

//         const tagsParts = [];
//         if (Number(emp.isDismissedWeek) === 1 && isDismissalOrAfter) tagsParts.push("DP");
//         if (hasP) tagsParts.push("P");
//         if (hasL) tagsParts.push("L");
//         if (hasE) tagsParts.push("E");
//         if (hasEP) tagsParts.push("EP");
//         if (hasT) tagsParts.push("T");
//         const tags = tagsParts.join(".") || "-";

//         const entryF = fmt12smart(entry, { bias: "entry" });
//         const exitF = fmt12smart(exit, { bias: "exit", refEntry: entry });

//         // ===== Cálculo H.T / H.A =====
//         let ht = "";
//         let ha = "";

//         let cfgToUse = shiftCfg;

//         if (isNightShift && marksInDay) {
//           const startMin = 7 * 60;
//           const endMin = 16 * 60 + 45;
//           const lunchHours = LUNCH_HOURS_DEFAULT;
//           const hrsLaborables = (endMin - startMin) / 60 - lunchHours;
//           const lunchThresholdMin = startMin + LUNCH_OFFSET_MIN;
//           cfgToUse = { startMin, endMin, lunchHours, lunchThresholdMin, hrsLaborables };
//         }

//         const hasEntry = !!entry;
//         const hasExit = !!exit;
//         const hasPerm = permList.length > 0;

//         if (!hasEntry && !hasExit && !hasPerm) {
//           ht = "";
//           ha = "";
//         } else if ((hasEntry && !hasExit) || (!hasEntry && hasExit)) {
//           ht = 0;
//           ha = 0;
//         } else {
//           const exceptionTime = exceptionTimeMap.get(String(emp.employeeID)) || 0;
//           const { hrsTrabajadas, hrsAusencia } = computeDaySummary(
//             entry,
//             exit,
//             permList,
//             exceptionTime,
//             cfgToUse
//           );

//           ht = hrsTrabajadas == null ? "" : Number(hrsTrabajadas.toFixed(2));
//           ha = hrsAusencia == null ? "" : Number(hrsAusencia.toFixed(2));
//         }

//         row.push(tags);
//         row.push(entryF);
//         row.push(exitF);
//         row.push(ht === "" ? "" : ht);
//         row.push(ha === "" ? "" : ha);

//         // ====== Permisos ======
//         const list = permList.slice(0, maxPermsByDay[dIdx]);
//         const firstSolicIdx = list.findIndex((p) => p.request === 1 && !p.exitFmt && !p.entryFmt);

//         for (let iPerm = 0; iPerm < maxPermsByDay[dIdx]; iPerm++) {
//           const p = list[iPerm];
//           if (!p) {
//             row.push("-", "-");
//             continue;
//           }

//           if (noAttendance && iPerm === firstSolicIdx) {
//             row.push(NO_SHOW, NO_SHOW);
//             weekComments.push(`${d.abbr} P${iPerm + 1}: No vino, no marco asistencia`);
//             continue;
//           }

//           if (p.request === 0 && !p.exitFmt && !p.entryFmt) {
//             const reason = p.comment ? ` – Justificación: ${p.comment}` : "";
//             row.push(`${NO_HOUR}${reason}`, `${NO_HOUR}${reason}`);
//             weekComments.push(`${d.abbr} P${iPerm + 1}: sin hora${p.comment ? ` (Justificación: ${p.comment})` : ""}`);
//             continue;
//           }

//           const sVal = fmt12smart(p.exitFmt, { bias: "auto" });
//           const eVal = fmt12smart(p.entryFmt, { bias: "auto" });
//           row.push(sVal, eVal);
//         }

//         if (includeDispatchByDay[dIdx]) {
//           row.push(fmt12smart(isAfterDismissalDay ? "" : (rec.dispatchingTime || ""), { bias: "auto" }));
//         }

//         if (!isAfterDismissalDay) {
//           if (rec.exitComment) weekComments.push(`${d.abbr} Salida: ${rec.exitComment}`);
//           if (rec.dispatchingComment) weekComments.push(`${d.abbr} Despacho: ${rec.dispatchingComment}`);
//         }
//       });

//       row.push(weekComments.join(" | ") || "");
//       return row;
//     });

//     tableRows.forEach((r) => ws.addRow(r));

//     ws.eachRow((row, rNum) => {
//       if (rNum > 4) {
//         row.eachCell((cell) => {
//           cell.font = { name: "Calibri", size: 11 };
//           cell.alignment = { horizontal: "center", vertical: "middle" };
//           cell.border = {
//             top: { style: "thin" },
//             bottom: { style: "thin" },
//             left: { style: "thin" },
//             right: { style: "thin" },
//           };
//         });
//         row.fill = {
//           type: "pattern",
//           pattern: "solid",
//           fgColor: { argb: rNum % 2 === 1 ? "F5F5F5" : "FFFFFF" },
//         };
//         row.height = 20;
//       }
//     });

//     // ====== Resaltar DP en rojo (celda de TAGS por día) ======
//     const DP_BG = "C00000";   // rojo fuerte
//     const DP_FONT = "FFFFFF"; // blanco

//     // Las filas de datos empiezan después del header (tú tienes 4 filas de header)
//     const firstDataRow = 5;

//     for (let r = firstDataRow; r <= ws.rowCount; r++) {
//       for (let dIdx = 0; dIdx < dayStart.length; dIdx++) {
//         const tagsCol = dayStart[dIdx]; // en tu layout, tags está en el inicio del bloque del día
//         const cell = ws.getCell(r, tagsCol);
//         const val = String(cell.value ?? "");

//         // Si en tags viene "DP" o "P.DP" o "DP.P" etc...
//         if (val.includes("DP")) {
//           cell.fill = {
//             type: "pattern",
//             pattern: "solid",
//             fgColor: { argb: DP_BG },
//           };
//           cell.font = {
//             name: "Calibri",
//             size: 11,
//             bold: true,
//             color: { argb: DP_FONT },
//           };
//           cell.alignment = { horizontal: "center", vertical: "middle" };
//         }
//       }
//     }

//     const widths = [{ width: 10 }, { width: 10 }, { width: 30 }, { width: 16 }];
//     days.forEach((_, i) => {
//       widths.push({ width: 12 }, { width: 12 }, { width: 12 }, { width: 10 }, { width: 10 });
//       for (let j = 1; j <= maxPermsByDay[i]; j++) {
//         widths.push({ width: 12 }, { width: 12 });
//       }
//       if (includeDispatchByDay[i]) widths.push({ width: 12 });
//     });
//     widths.push({ width: 60 });
//     ws.columns = widths;

//     const buf = await workbook.xlsx.writeBuffer();
//     const filename = `asistencia_semanal_${isActive == 0 ? "inactivos" : "activos"}_semana${selectedWeek}_${year}.xlsx`;
//     res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
//     res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
//     res.send(buf);
//   } catch (err) {
//     console.error("Error al exportar asistencia semanal:", err.stack || err.message);
//     res.status(500).send({
//       message: `Error interno del servidor al generar el archivo Excel: ${err.message}`,
//     });
//   }
// };

exports.exportWeeklyAttendance = async (req, res) => {
  try {
    const { weeklyAttendance, selectedMonth, selectedWeek, isActive } = req.body;

    //  Permitimos que weeklyAttendance venga vacío, porque ahora completamos desde BD
    if (!weeklyAttendance || !Array.isArray(weeklyAttendance))
      throw new Error("weeklyAttendance is missing or invalid.");
    if (!selectedMonth || typeof selectedMonth !== "string")
      throw new Error("selectedMonth is missing or invalid.");
    if (!selectedWeek || typeof selectedWeek !== "string")
      throw new Error("selectedWeek is missing or invalid.");

    // ====== Helpers de tiempo ======
    const parseTime = (t) =>
      dayjs(t, ["HH:mm:ss", "H:mm:ss", "hh:mm:ss A", "h:mm:ss A"], true);

    const fmt12smart = (t, { bias = "auto", refEntry } = {}) => {
      if (!t) return "-";
      let d = dayjs(
        t,
        [
          "HH:mm:ss",
          "H:mm:ss",
          "hh:mm:ss A",
          "h:mm:ss A",
          "YYYY-MM-DD HH:mm:ss",
          "YYYY-MM-DD hh:mm:ss A",
        ],
        true
      );
      if (!d.isValid()) return String(t);

      const raw = String(t);
      const hasMeridiem = /AM|PM/i.test(raw);
      if (!hasMeridiem && bias === "exit" && refEntry) {
        const e = dayjs(
          refEntry,
          ["HH:mm:ss", "H:mm:ss", "hh:mm:ss A", "h:mm:ss A"],
          true
        );
        if (e.isValid()) {
          const hrEntry = e.hour();
          const hrExit = d.hour();
          if (hrEntry >= 5 && hrEntry <= 9 && hrExit >= 1 && hrExit <= 6)
            d = d.add(12, "hour");
        }
      }
      return d.format("hh:mm:ss A");
    };

    const timeToMinutes = (t) => {
      if (!t) return null;
      const d = parseTime(t);
      if (!d.isValid()) return null;
      return d.hour() * 60 + d.minute();
    };

    const mergeIntervals = (intervals) => {
      if (!intervals.length) return [];
      const sorted = intervals
        .slice()
        .sort((a, b) => a[0] - b[0] || a[1] - b[1]);
      const result = [];
      let [curStart, curEnd] = sorted[0];
      for (let i = 1; i < sorted.length; i++) {
        const [s, e] = sorted[i];
        if (s <= curEnd) {
          if (e > curEnd) curEnd = e;
        } else {
          result.push([curStart, curEnd]);
          curStart = s;
          curEnd = e;
        }
      }
      result.push([curStart, curEnd]);
      return result;
    };

    const totalMinutes = (intervals) => {
      const merged = mergeIntervals(intervals);
      return merged.reduce((sum, [s, e]) => sum + (e - s), 0);
    };

    const coverageInSegment = (intervals, segStart, segEnd) => {
      if (segStart == null || segEnd == null || segEnd <= segStart) return 0;
      const clipped = [];
      for (const [s, e] of intervals) {
        const cs = Math.max(segStart, s);
        const ce = Math.min(segEnd, e);
        if (ce > cs) clipped.push([cs, ce]);
      }
      return totalMinutes(clipped);
    };

    // ====== Constantes / colores ======
    const TITLE_BG = "1F3864";
    const WHITE = "FFFFFF";
    const GREEN_E = "C6EFCE";
    const RED_S = "FFC7CE";
    const GRAY_P = "D6DCE4";
    const YELLOW_D = "FFEB9C";
    const NO_HOUR = "(sin hora)";
    const NO_SHOW = "(No vino, no marco asistencia)";

    // Almuerzo genérico
    const LUNCH_HOURS_DEFAULT = 0.75; // 45 minutos
    const LUNCH_OFFSET_MIN = 5.75 * 60; // 5h45 después de la entrada

    const dayHeaderColors = {
      Lunes: "FFFFFF",
      Martes: "FFC0CB",
      Miércoles: "FFFFFF",
      Jueves: "FFFFFF",
      Viernes: "FFFFFF",
      Sábado: "D3D3D3",
      Domingo: "D3D3D3",
    };

    // ====== Semana ISO (ANTES de buscar empleados) ======
    let startOfWeek;
    if (weeklyAttendance.length > 0) {
      const firstDate = weeklyAttendance.map((r) => String(r.date)).sort()[0];
      startOfWeek = dayjs(firstDate).startOf("isoWeek");
    } else {
      const yearNow = dayjs().year();
      startOfWeek = dayjs()
        .year(yearNow)
        .isoWeek(parseInt(selectedWeek))
        .startOf("isoWeek");
    }
    const endOfWeek = startOfWeek.add(6, "day");

    const days = Array.from({ length: 7 }, (_, i) => {
      const d = startOfWeek.add(i, "day");
      return {
        idx: i,
        iso: d.isoWeekday(), // 1 = lunes ... 7 = domingo
        date: d.format("YYYY-MM-DD"),
        name:
          d.format("dddd").charAt(0).toUpperCase() + d.format("dddd").slice(1),
        short: d.format("DD/MM"),
        abbr: d.format("ddd"),
      };
    });

    // ====== Empleados según isActive + despedidos dentro de la semana ======
    const [activeEmployees] = await db.query(
      `
      SELECT 
        e.employeeID,
        e.shiftID,
        CONCAT(e.firstName,' ',COALESCE(e.middleName,''),' ',e.lastName) AS employeeName,
        CASE WHEN d.employeeID IS NOT NULL THEN 1 ELSE 0 END AS isDismissedWeek,
        d.dateDismissal
      FROM employees_emp e
      LEFT JOIN (
        SELECT hd.employeeID, MAX(DATE(hd.dateDismissal)) AS dateDismissal
        FROM h_dismissal_emp hd
        WHERE DATE(hd.dateDismissal) BETWEEN ? AND ?
        GROUP BY hd.employeeID
      ) d ON d.employeeID = e.employeeID
      WHERE e.isActive = ?
         OR d.employeeID IS NOT NULL
      ORDER BY e.employeeID
      `,
      [
        startOfWeek.format("YYYY-MM-DD"),
        endOfWeek.format("YYYY-MM-DD"),
        Number(isActive),
      ]
    );

    const activeSet = new Set(activeEmployees.map((e) => String(e.employeeID)));

    // =====================================================================
    // ✅ CLAVE: NO depender del frontend
    // Traemos marcajes reales desde BD para TODOS los empleados del reporte
    // =====================================================================
    const ids = [...activeSet];
    let weeklyAttendanceFinal = Array.isArray(weeklyAttendance)
      ? weeklyAttendance.slice()
      : [];

    if (ids.length > 0) {
      // 1) asistencia real (h_attendance_emp)
      const [attRows] = await db.query(
        `
        SELECT 
          a.employeeID,
          DATE(a.date) AS date,
          TIME_FORMAT(a.entryTime,'%H:%i:%s') AS entryTime,
          TIME_FORMAT(a.exitTime, '%H:%i:%s') AS exitTime,
          COALESCE(a.comment,'') AS exitComment
        FROM h_attendance_emp a
        WHERE DATE(a.date) BETWEEN ? AND ?
          AND a.employeeID IN (${ids.map(() => "?").join(",")})
        `,
        [startOfWeek.format("YYYY-MM-DD"), endOfWeek.format("YYYY-MM-DD"), ...ids]
      );

      // 2) despacho real (dispatching_emp)
      const [dispRows] = await db.query(
        `
        SELECT
          d.employeeID,
          DATE(d.date) AS date,
          TIME_FORMAT(d.exitTimeComplete,'%H:%i:%s') AS dispatchingTime,
          COALESCE(d.comment,'') AS dispatchingComment
        FROM dispatching_emp d
        WHERE DATE(d.date) BETWEEN ? AND ?
          AND d.employeeID IN (${ids.map(() => "?").join(",")})
        `,
        [startOfWeek.format("YYYY-MM-DD"), endOfWeek.format("YYYY-MM-DD"), ...ids]
      );

      const keyOf = (empId, dateStr) =>
        `${String(empId)}|${dayjs(String(dateStr)).format("YYYY-MM-DD")}`;

      const merged = new Map();

      // A) cargar lo que venga del frontend
      for (const r of weeklyAttendanceFinal) {
        const key = keyOf(r.employeeID, r.date);
        merged.set(key, {
          employeeID: String(r.employeeID),
          date: dayjs(String(r.date)).format("YYYY-MM-DD"),
          entryTime: String(r.entryTime ?? ""),
          exitTime: String(r.exitTime ?? ""),
          dispatchingTime: String(r.dispatchingTime ?? ""),
          dispatchingComment: String(r.dispatchingComment ?? ""),
          exitComment: String(r.exitComment ?? ""),
        });
      }

      // B) completar con asistencia de BD
      for (const r of attRows || []) {
        const key = keyOf(r.employeeID, r.date);
        const prev = merged.get(key) || {
          employeeID: String(r.employeeID),
          date: dayjs(String(r.date)).format("YYYY-MM-DD"),
          entryTime: "",
          exitTime: "",
          dispatchingTime: "",
          dispatchingComment: "",
          exitComment: "",
        };

        merged.set(key, {
          ...prev,
          entryTime: prev.entryTime || (r.entryTime || ""),
          exitTime: prev.exitTime || (r.exitTime || ""),
          exitComment: prev.exitComment || (r.exitComment || ""),
        });
      }

      // C) completar con despacho de BD
      for (const r of dispRows || []) {
        const key = keyOf(r.employeeID, r.date);
        const prev = merged.get(key) || {
          employeeID: String(r.employeeID),
          date: dayjs(String(r.date)).format("YYYY-MM-DD"),
          entryTime: "",
          exitTime: "",
          dispatchingTime: "",
          dispatchingComment: "",
          exitComment: "",
        };

        merged.set(key, {
          ...prev,
          dispatchingTime: prev.dispatchingTime || (r.dispatchingTime || ""),
          dispatchingComment:
            prev.dispatchingComment || (r.dispatchingComment || ""),
        });
      }

      weeklyAttendanceFinal = Array.from(merged.values());
    }

    // ====== Mapa de Tipo de Planilla ======
    const [payrollRows] = await db.query(`
      SELECT e.employeeID, COALESCE(pt.payrollName,'') AS payrollName
      FROM employees_emp e
      LEFT JOIN payrolltype_emp pt ON pt.payrollTypeID = e.payrollTypeID
    `);
    const payrollMap = new Map(
      payrollRows.map((r) => [String(r.employeeID), r.payrollName || ""])
    );

    // ====== Detalle de turnos por día (detailshift_emp) ======
    const [shiftDetailRows] = await db.query(`
      SELECT shiftID, day, startTime, endTime
      FROM detailsshift_emp
    `);

    const dayNameToIso = {
      Monday: 1,
      Tuesday: 2,
      Wednesday: 3,
      Thursday: 4,
      Friday: 5,
      Saturday: 6,
      Sunday: 7,
    };

    const shiftDetailMap = new Map();
    for (const r of shiftDetailRows) {
      const isoDay = dayNameToIso[r.day];
      if (!isoDay) continue;
      let s = timeToMinutes(r.startTime);
      let e = timeToMinutes(r.endTime);
      if (s == null || e == null) continue;
      if (e <= s) e += 24 * 60;
      shiftDetailMap.set(`${r.shiftID}|${isoDay}`, { startMin: s, endMin: e });
    }

    const getShiftConfig = (shiftId, isoDay) => {
      const key = `${shiftId}|${isoDay}`;
      const base = shiftDetailMap.get(key);

      if (base) {
        const { startMin, endMin } = base;
        const lunchHours = LUNCH_HOURS_DEFAULT;
        const hrsLaborables = (endMin - startMin) / 60 - lunchHours;
        const lunchThresholdMin = startMin + LUNCH_OFFSET_MIN;
        return { startMin, endMin, lunchHours, lunchThresholdMin, hrsLaborables };
      }

      const startMin = 7 * 60;
      const endMin = 16 * 60 + 45;
      const lunchHours = LUNCH_HOURS_DEFAULT;
      const hrsLaborables = (endMin - startMin) / 60 - lunchHours;
      const lunchThresholdMin = startMin + LUNCH_OFFSET_MIN;
      return { startMin, endMin, lunchHours, lunchThresholdMin, hrsLaborables };
    };

    // ====== Mapa de exceptionTime ======
    const [excRows] = await db.query(`
      SELECT e.employeeID, COALESCE(ex.exceptionTime, 0) AS exceptionTime
      FROM employees_emp e
      LEFT JOIN exceptions_emp ex ON ex.exceptionID = e.exceptionID
      WHERE e.employeeID IN (${[...activeSet].map((id) => db.escape(id)).join(",")})
    `);
    const exceptionTimeMap = new Map(
      excRows.map((r) => [String(r.employeeID), Number(r.exceptionTime || 0)])
    );

    // ====== Indexar asistencias (YA con weeklyAttendanceFinal) ======
    const attMap = new Map(
      weeklyAttendanceFinal
        .filter((r) => activeSet.has(String(r.employeeID)))
        .map((r) => {
          const key = `${String(r.employeeID)}|${String(r.date)}`;
          return [
            key,
            {
              entryTime: String(r.entryTime ?? ""),
              exitTime: String(r.exitTime ?? ""),
              dispatchingTime: String(r.dispatchingTime ?? ""),
              dispatchingComment: String(r.dispatchingComment ?? ""),
              exitComment: String(r.exitComment ?? ""),
            },
          ];
        })
    );

    // ====== Grid base: todos los empleados × 7 días ======
    const san = [];
    for (const emp of activeEmployees) {
      const empId = String(emp.employeeID);
      for (const d of days) {
        const key = `${empId}|${d.date}`;
        const rec =
          attMap.get(key) || {
            entryTime: "",
            exitTime: "",
            dispatchingTime: "",
            dispatchingComment: "",
            exitComment: "",
          };
        san.push({
          employeeID: empId,
          employeeName: String(emp.employeeName ?? ""),
          date: d.date,
          ...rec,
        });
      }
    }

    // ====== Permisos (aprobados) de la semana ======
    const [permWeekRows] = await db.query(
      `
      SELECT employeeID,
             DATE(date) AS date,
             TIME_FORMAT(exitPermission,  '%H:%i:%s') AS exitFmt,
             TIME_FORMAT(entryPermission, '%H:%i:%s') AS entryFmt,
             request,
             comment,
             lunchTime,
             createdDate,
             permissionID
      FROM permissionattendance_emp
      WHERE date BETWEEN ? AND ?
        AND isApproved = 1
      ORDER BY employeeID, date, createdDate, permissionID
      `,
      [startOfWeek.format("YYYY-MM-DD"), endOfWeek.format("YYYY-MM-DD")]
    );

    const permsByEmpDate = new Map();
    for (const p of permWeekRows) {
      const emp = String(p.employeeID);
      if (!activeSet.has(emp)) continue;
      const key = `${emp}|${dayjs(p.date).format("YYYY-MM-DD")}`;
      if (!permsByEmpDate.has(key)) permsByEmpDate.set(key, []);
      permsByEmpDate.get(key).push({
        request: p.request == 0 || p.request === "0" ? 0 : 1,
        exitFmt: p.exitFmt || "",
        entryFmt: p.entryFmt || "",
        comment: p.comment || "",
        lunchTime: p.lunchTime,
      });
    }

    // ====== Excepciones por día (L, E, EP) ======
    let hasBridge = false;
    let bridgeRows = [];
    try {
      const [rowsBridge] = await db.query(`
        SELECT 
          ee.employeeID,
          ee.isActive,
          ee.startDate,
          ee.endDate,
          UPPER(ex.exceptionName) AS exceptionName
        FROM employee_exceptions ee
        JOIN exceptions_emp ex ON ex.exceptionID = ee.exceptionID
        WHERE ee.employeeID IN (${[...activeSet].map((id) => db.escape(id)).join(",")})
      `);
      bridgeRows = rowsBridge || [];
      hasBridge = bridgeRows.length > 0;
    } catch (_) {
      hasBridge = false;
      bridgeRows = [];
    }

    const flagsByEmpDate = new Map();
    if (hasBridge) {
      for (const empId of activeSet) {
        const myExc = bridgeRows.filter(
          (r) => String(r.employeeID) === String(empId)
        );
        for (const d of days) {
          const applies = (rec) => {
            const active =
              rec.isActive == null || rec.isActive === 1 || rec.isActive === "1";
            const sdOk =
              !rec.startDate ||
              dayjs(d.date).isSameOrAfter(dayjs(rec.startDate), "day");
            const edOk =
              !rec.endDate ||
              dayjs(d.date).isSameOrBefore(dayjs(rec.endDate), "day");
            return active && sdOk && edOk;
          };
          const norm = (x) => (x || "").toString().trim().toUpperCase();
          const L = myExc.some(
            (r) => applies(r) && norm(r.exceptionName) === "LACTANCIA"
          )
            ? 1
            : 0;
          const E = myExc.some(
            (r) => applies(r) && norm(r.exceptionName) === "EMBARAZADA"
          )
            ? 1
            : 0;
          const EP = myExc.some(
            (r) => applies(r) && norm(r.exceptionName) === "PERMISO ESPECIAL"
          )
            ? 1
            : 0;
          flagsByEmpDate.set(`${empId}|${d.date}`, { L, E, EP });
        }
      }
    } else {
      const [fkRows] = await db.query(`
        SELECT 
          e.employeeID,
          UPPER(COALESCE(ex.exceptionName,'')) AS exName
        FROM employees_emp e
        LEFT JOIN exceptions_emp ex ON ex.exceptionID = e.exceptionID
        WHERE e.employeeID IN (${[...activeSet].map((id) => db.escape(id)).join(",")})
      `);
      const fkMap = new Map(
        fkRows.map((r) => [String(r.employeeID), r.exName || ""])
      );
      for (const empId of activeSet) {
        const exName = (fkMap.get(String(empId)) || "").toUpperCase();
        const L = exName === "LACTANCIA" ? 1 : 0;
        const E = exName === "EMBARAZADA" ? 1 : 0;
        const EP = 0;
        for (const d of days)
          flagsByEmpDate.set(`${empId}|${d.date}`, { L, E, EP });
      }
    }

    // =====================================================================
    // ✅ INCAPACIDAD / MATERNIDAD POR DÍA (palabra "Incapacidad")
    // - No oculta marcajes
    // - Marca si el rango cubre el día
    // =====================================================================
    const incapByEmpDate = new Map();
    try {
      const [incRows] = await db.query(
        `
        SELECT employeeID, kind, startDate, endDate
        FROM (
          SELECT
            d.employeeID,
            'DIS' AS kind,
            DATE(d.startDate) AS startDate,
            DATE(d.endDate) AS endDate
          FROM disability_emp d
          WHERE DATE(d.startDate) <= ? AND DATE(d.endDate) >= ?

          UNION ALL

          SELECT
            m.employeeID,
            'MAT' AS kind,
            DATE(m.startDate) AS startDate,
            DATE(m.endDate) AS endDate
          FROM maternity_emp m
          WHERE DATE(m.startDate) <= ? AND DATE(m.endDate) >= ?
        ) x
        `,
        [
          endOfWeek.format("YYYY-MM-DD"),
          startOfWeek.format("YYYY-MM-DD"),
          endOfWeek.format("YYYY-MM-DD"),
          startOfWeek.format("YYYY-MM-DD"),
        ]
      );

      for (const r of incRows || []) {
        const empId = String(r.employeeID);
        const sd = dayjs(r.startDate).startOf("day");
        const ed = dayjs(r.endDate).startOf("day");
        for (const d of days) {
          const dd = dayjs(d.date).startOf("day");
          if (dd.isSameOrAfter(sd) && dd.isSameOrBefore(ed)) {
            const key = `${empId}|${d.date}`;
            // guardamos solo un flag (si querés diferenciar DIS vs MAT, aquí se puede)
            incapByEmpDate.set(key, { kind: r.kind });
          }
        }
      }
    } catch (e) {
      // si algo falla, NO rompemos el export (solo no marca Incapacidad)
      console.error("No se pudo cargar incapacidad/maternidad:", e.message);
    }

    // ====== Función para cálculo diario (H.T / H.A) ======
    const computeDaySummary = (entry, exit, permsList, exceptionTime = 0, shiftCfg) => {
      const { startMin, endMin, lunchHours, lunchThresholdMin, hrsLaborables } = shiftCfg;
      const crossesMidnight = endMin > 24 * 60;

      const adjustForShift = (min) => {
        if (min == null) return null;
        if (crossesMidnight && min < startMin && min <= 12 * 60) min += 24 * 60;
        return min;
      };

      let entryMin = entry ? adjustForShift(timeToMinutes(entry)) : null;
      let exitMin = exit ? adjustForShift(timeToMinutes(exit)) : null;

      const entryMinClamped =
        entryMin == null ? null : Math.min(Math.max(entryMin, startMin), endMin);
      const exitMinClamped =
        exitMin == null ? null : Math.min(Math.max(exitMin, startMin), endMin);

      const permisoIntervalsRaw = [];
      let permisoLunchFlag = false;

      for (const p of permsList) {
        let s = adjustForShift(timeToMinutes(p.exitFmt));
        let e = adjustForShift(timeToMinutes(p.entryFmt));
        if (s == null || e == null) continue;
        if (e <= s) continue;
        permisoIntervalsRaw.push([s, e]);
        if (Number(p.lunchTime || 0) === 1) permisoLunchFlag = true;
      }

      const permisoMinInShift = coverageInSegment(permisoIntervalsRaw, startMin, endMin);
      let hrsPermiso = permisoMinInShift / 60.0;
      if (permisoLunchFlag && hrsPermiso > 0) hrsPermiso = Math.max(hrsPermiso - lunchHours, 0);

      const anyLunchFlag = permisoLunchFlag;
      const soloEntradaSinSalida = !!entry && !exit;
      const soloSalidaSinEntrada = !entry && !!exit;
      const caso3 = soloEntradaSinSalida || soloSalidaSinEntrada;

      let hrsAusente_entrada = 0;
      let hrsAusente_salida = 0;
      let hrsAusencia = 0;
      let hrsTrabajadas = 0;

      if (caso3) {
        hrsTrabajadas = 0;
        hrsAusencia = 0;
      } else {
        const noEntry = !entry;
        const noExit = !exit;

        if (noEntry && noExit && permisoMinInShift === 0) {
          const netPermiso = hrsPermiso;
          hrsAusencia = Math.max(hrsLaborables - netPermiso, 0);
        } else {
          if (entryMinClamped != null && entryMinClamped > startMin) {
            const totalGapMin = entryMinClamped - startMin;
            const coveredMin = coverageInSegment(permisoIntervalsRaw, startMin, entryMinClamped);
            let baseMin = totalGapMin - coveredMin;
            if (baseMin < 0) baseMin = 0;

            if (baseMin > 0 && entryMinClamped >= lunchThresholdMin && !anyLunchFlag) {
              baseMin -= lunchHours * 60;
              if (baseMin < 0) baseMin = 0;
            }

            hrsAusente_entrada = baseMin / 60.0;
          }

          if (exitMinClamped != null && exitMinClamped < endMin) {
            const totalGapMin = endMin - exitMinClamped;
            const coveredAfterExit = coverageInSegment(permisoIntervalsRaw, exitMinClamped, endMin);
            let effectiveGapMin = totalGapMin - coveredAfterExit;
            if (effectiveGapMin < 0) effectiveGapMin = 0;

            let baseHours = effectiveGapMin / 60.0 - Number(exceptionTime || 0);
            if (baseHours < 0) baseHours = 0;

            hrsAusente_salida = baseHours;
          }

          hrsAusencia = hrsAusente_entrada + hrsAusente_salida;
        }

        hrsTrabajadas = Math.max(hrsLaborables - (hrsPermiso + hrsAusencia), 0);
      }

      return { hrsTrabajadas, hrsAusencia };
    };

    // ====== Máximo de permisos por día ======
    const maxPermsByDay = Array(7).fill(0);
    for (const [key, list] of permsByEmpDate.entries()) {
      const [, dateStr] = key.split("|");
      const idx = dayjs(dateStr).diff(startOfWeek, "day");
      if (idx < 0 || idx > 6) continue;
      maxPermsByDay[idx] = Math.max(maxPermsByDay[idx], Math.min(list.length, 5));
    }

    // ====== ¿Hay despacho por día? ======
    const includeDispatchByDay = Array(7).fill(false);
    for (let i = 0; i < 7; i++) {
      includeDispatchByDay[i] = san.some(
        (r) => r.date === days[i].date && r.dispatchingTime && activeSet.has(r.employeeID)
      );
    }

    // ====== Si no hay empleados ======
    if (activeEmployees.length === 0) {
      const workbook0 = new ExcelJS.Workbook();
      const w = workbook0.addWorksheet("Asistencia Semanal");
      w.addRow([`Reporte Semanal - ${isActive == 0 ? "INACTIVOS" : "ACTIVOS"} - (sin empleados)`]);
      w.mergeCells(1, 1, 1, 5);
      w.getRow(1).font = { bold: true };
      w.addRow(["Item", "Código", "Empleado", "Tipo de Planilla", "Comentarios"]);
      w.columns = [{ width: 10 }, { width: 10 }, { width: 30 }, { width: 18 }, { width: 60 }];
      const buf0 = await workbook0.xlsx.writeBuffer();
      const filename0 = `asistencia_semanal_${isActive == 0 ? "inactivos" : "activos"}_semana${selectedWeek}_${dayjs().year()}.xlsx`;
      res.setHeader("Content-Disposition", `attachment; filename="${filename0}"`);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      return res.send(buf0);
    }

    // ====== Excel (estructura con H.T / H.A + permisos + despacho) ======
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("Asistencia Semanal");

    const year = dayjs().year();
    const monthName = dayjs().month(parseInt(selectedMonth)).format("MMMM").toUpperCase();

    const columnsPerDay = days.map(
      (_, i) => 1 + 4 + maxPermsByDay[i] * 2 + (includeDispatchByDay[i] ? 1 : 0)
    );
    const totalDataCols = 4 + columnsPerDay.reduce((a, b) => a + b, 0) + 1;

    const title = ws.addRow([`Reporte Semanal  - Mes ${monthName} Semana ${selectedWeek}`]);
    ws.mergeCells(1, 1, 1, totalDataCols);
    title.font = { name: "Calibri", size: 16, bold: true, color: { argb: WHITE } };
    title.alignment = { horizontal: "center", vertical: "middle" };
    title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: TITLE_BG } };
    ws.getRow(1).height = 30;

    // ✅ SUBTITLE con palabras (ya no siglas)
    const subtitle = ws.addRow([
      `Total de empleados: ${activeEmployees.length}  |  ${
        isActive == 0 ? "INACTIVOS" : "ACTIVOS"
      } | Permiso / Lactancia / Embarazada / Permiso Especial / Tarde / Despido / Incapacidad`,
    ]);
    ws.mergeCells(2, 1, 2, totalDataCols);
    subtitle.font = { name: "Calibri", size: 12, bold: true };
    subtitle.alignment = { horizontal: "center", vertical: "middle" };
    subtitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "F2F2F2" } };
    ws.getRow(2).height = 24;

    const mainHeaderRow = [];
    mainHeaderRow.push("", "", "", "");
    columnsPerDay.forEach((n) => mainHeaderRow.push(...Array(n).fill("")));
    mainHeaderRow.push("");
    ws.addRow(mainHeaderRow);

    // ✅ header de tags con palabras
    const subHeaderRow = ["Item", "Código", "Empleado", "Tipo de Planilla"];
    days.forEach((_, i) => {
      subHeaderRow.push("(Estados)", "E", "S", "H.T", "H.A");
      for (let j = 1; j <= maxPermsByDay[i]; j++) subHeaderRow.push(`P${j}S`, `P${j}E`);
      if (includeDispatchByDay[i]) subHeaderRow.push("D");
    });
    subHeaderRow.push("Comentarios");
    ws.addRow(subHeaderRow);

    ws.mergeCells(3, 1, 4, 1);
    ws.mergeCells(3, 2, 4, 2);
    ws.mergeCells(3, 3, 4, 3);
    ws.mergeCells(3, 4, 4, 4);
    ws.mergeCells(3, totalDataCols, 4, totalDataCols);

    const setStatic = (col, label) => {
      const cell = ws.getCell(3, col);
      cell.value = label;
      cell.font = { name: "Calibri", size: 12, bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: WHITE } };
      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
      };
    };
    setStatic(1, "Item");
    setStatic(2, "Código");
    setStatic(3, "Empleado");
    setStatic(4, "Tipo de Planilla");
    setStatic(totalDataCols, "Comentarios");

    let cur = 5;
    const dayStart = [];
    days.forEach((d) => {
      const span = columnsPerDay[d.idx];
      dayStart.push(cur);
      ws.mergeCells(3, cur, 3, cur + span - 1);
      const c = ws.getCell(3, cur);
      c.value = `${d.name} ${d.short}`;
      c.font = { name: "Calibri", size: 12, bold: true };
      c.alignment = { horizontal: "center", vertical: "middle" };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: dayHeaderColors[d.name] || WHITE } };
      c.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
      };
      cur += span;
    });
    ws.getRow(3).height = 20;

    ws.getRow(4).font = { name: "Calibri", size: 11, bold: true };
    ws.getRow(4).alignment = { horizontal: "center", vertical: "middle" };
    ws.getRow(4).eachCell((cell) => {
      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
      };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: WHITE } };
    });

    days.forEach((_, i) => {
      const s = dayStart[i];
      ws.getRow(4).getCell(s + 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: GREEN_E } };
      ws.getRow(4).getCell(s + 2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: RED_S } };
      const m = maxPermsByDay[i];
      for (let j = 0; j < m; j++) {
        ws.getRow(4).getCell(s + 5 + j * 2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRAY_P } };
        ws.getRow(4).getCell(s + 5 + j * 2 + 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRAY_P } };
      }
      if (includeDispatchByDay[i]) {
        ws.getRow(4).getCell(s + 5 + m * 2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: YELLOW_D } };
      }
    });

    // ====== Filas de datos ======
    const tableRows = activeEmployees.map((emp, idx) => {
      const row = [idx + 1, emp.employeeID, emp.employeeName, payrollMap.get(String(emp.employeeID)) || ""];
      const weekComments = [];

      days.forEach((d, dIdx) => {
        const rec =
          san.find((r) => r.employeeID === String(emp.employeeID) && r.date === d.date) || {};

        const dismissedDate = emp.dateDismissal ? dayjs(emp.dateDismissal).format("YYYY-MM-DD") : null;

        // ✅ Regla final:
        // - Antes del despido: normal
        // - Día del despido: normal + Despido
        // - Después del despido: limpiar horas/permisos/despacho, pero Despido se mantiene
        const isAfterDismissalDay =
          dismissedDate && dayjs(d.date).isAfter(dayjs(dismissedDate), "day");

        const isDismissalOrAfter =
          dismissedDate &&
          (dayjs(d.date).isSame(dayjs(dismissedDate), "day") ||
            dayjs(d.date).isAfter(dayjs(dismissedDate), "day"));

        const entry = isAfterDismissalDay ? "" : (rec.entryTime || "");
        const exit  = isAfterDismissalDay ? "" : (rec.exitTime || "");
        const noAttendance = !entry && !exit;

        const permList = isAfterDismissalDay
          ? []
          : (permsByEmpDate.get(`${emp.employeeID}|${d.date}`) || []);

        const shiftCfg = getShiftConfig(emp.shiftID, d.iso);

        const isNightShift = shiftCfg.endMin > 24 * 60;
        const entryMinRaw = timeToMinutes(entry);
        const exitMinRaw = timeToMinutes(exit);

        const marksInDay =
          entryMinRaw != null &&
          exitMinRaw != null &&
          entryMinRaw >= 4 * 60 &&
          exitMinRaw <= 20 * 60 &&
          exitMinRaw > entryMinRaw;

        // ====== ESTADOS (ya NO siglas) ======
        const hasPermiso = permList.length > 0;

        const flags = flagsByEmpDate.get(`${emp.employeeID}|${d.date}`) || { L: 0, E: 0, EP: 0 };
        const hasLactancia = flags.L === 1;
        const hasEmbarazada = flags.E === 1;
        const hasPermisoEspecial = flags.EP === 1;

        // Incapacidad/Maternidad activa ese día
        const hasIncapacidad = !isAfterDismissalDay && incapByEmpDate.has(`${String(emp.employeeID)}|${d.date}`);

        let hasTarde = false;
        let startMinRef = shiftCfg.startMin;

        if (isNightShift && marksInDay) startMinRef = 7 * 60;
        if (entryMinRaw != null) hasTarde = entryMinRaw > startMinRef + 1;

        const estados = [];
        if (Number(emp.isDismissedWeek) === 1 && isDismissalOrAfter) estados.push("Despido");
        if (hasPermiso) estados.push("Permiso");
        if (hasLactancia) estados.push("Lactancia");
        if (hasEmbarazada) estados.push("Embarazada");
        if (hasPermisoEspecial) estados.push("Permiso Especial");
        if (hasTarde) estados.push("Tarde");
        if (hasIncapacidad) estados.push("Incapacidad");

        const tags = estados.join(" • ") || "-";

        const entryF = fmt12smart(entry, { bias: "entry" });
        const exitF = fmt12smart(exit, { bias: "exit", refEntry: entry });

        // ===== Cálculo H.T / H.A =====
        let ht = "";
        let ha = "";

        let cfgToUse = shiftCfg;

        if (isNightShift && marksInDay) {
          const startMin = 7 * 60;
          const endMin = 16 * 60 + 45;
          const lunchHours = LUNCH_HOURS_DEFAULT;
          const hrsLaborables = (endMin - startMin) / 60 - lunchHours;
          const lunchThresholdMin = startMin + LUNCH_OFFSET_MIN;
          cfgToUse = { startMin, endMin, lunchHours, lunchThresholdMin, hrsLaborables };
        }

        const hasEntry = !!entry;
        const hasExit = !!exit;
        const hasPerm = permList.length > 0;

        if (!hasEntry && !hasExit && !hasPerm) {
          ht = "";
          ha = "";
        } else if ((hasEntry && !hasExit) || (!hasEntry && hasExit)) {
          ht = 0;
          ha = 0;
        } else {
          const exceptionTime = exceptionTimeMap.get(String(emp.employeeID)) || 0;
          const { hrsTrabajadas, hrsAusencia } = computeDaySummary(
            entry,
            exit,
            permList,
            exceptionTime,
            cfgToUse
          );

          ht = hrsTrabajadas == null ? "" : Number(hrsTrabajadas.toFixed(2));
          ha = hrsAusencia == null ? "" : Number(hrsAusencia.toFixed(2));
        }

        row.push(tags);
        row.push(entryF);
        row.push(exitF);
        row.push(ht === "" ? "" : ht);
        row.push(ha === "" ? "" : ha);

        // ====== Permisos ======
        const list = permList.slice(0, maxPermsByDay[dIdx]);
        const firstSolicIdx = list.findIndex((p) => p.request === 1 && !p.exitFmt && !p.entryFmt);

        for (let iPerm = 0; iPerm < maxPermsByDay[dIdx]; iPerm++) {
          const p = list[iPerm];
          if (!p) {
            row.push("-", "-");
            continue;
          }

          if (noAttendance && iPerm === firstSolicIdx) {
            row.push(NO_SHOW, NO_SHOW);
            weekComments.push(`${d.abbr} P${iPerm + 1}: No vino, no marco asistencia`);
            continue;
          }

          if (p.request === 0 && !p.exitFmt && !p.entryFmt) {
            const reason = p.comment ? ` – Justificación: ${p.comment}` : "";
            row.push(`${NO_HOUR}${reason}`, `${NO_HOUR}${reason}`);
            weekComments.push(`${d.abbr} P${iPerm + 1}: sin hora${p.comment ? ` (Justificación: ${p.comment})` : ""}`);
            continue;
          }

          const sVal = fmt12smart(p.exitFmt, { bias: "auto" });
          const eVal = fmt12smart(p.entryFmt, { bias: "auto" });
          row.push(sVal, eVal);
        }

        if (includeDispatchByDay[dIdx]) {
          row.push(fmt12smart(isAfterDismissalDay ? "" : (rec.dispatchingTime || ""), { bias: "auto" }));
        }

        if (!isAfterDismissalDay) {
          if (rec.exitComment) weekComments.push(`${d.abbr} Salida: ${rec.exitComment}`);
          if (rec.dispatchingComment) weekComments.push(`${d.abbr} Despacho: ${rec.dispatchingComment}`);
        }
      });

      row.push(weekComments.join(" | ") || "");
      return row;
    });

    tableRows.forEach((r) => ws.addRow(r));

    ws.eachRow((row, rNum) => {
      if (rNum > 4) {
        row.eachCell((cell) => {
          cell.font = { name: "Calibri", size: 11 };
          cell.alignment = { horizontal: "center", vertical: "middle" };
          cell.border = {
            top: { style: "thin" },
            bottom: { style: "thin" },
            left: { style: "thin" },
            right: { style: "thin" },
          };
        });
        row.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: rNum % 2 === 1 ? "F5F5F5" : "FFFFFF" },
        };
        row.height = 20;
      }
    });

    // ====== Resaltar "Despido" en rojo (celda de Estados por día) ======
    const DP_BG = "C00000";   // rojo fuerte
    const DP_FONT = "FFFFFF"; // blanco

    const firstDataRow = 5;

    for (let r = firstDataRow; r <= ws.rowCount; r++) {
      for (let dIdx = 0; dIdx < dayStart.length; dIdx++) {
        const tagsCol = dayStart[dIdx];
        const cell = ws.getCell(r, tagsCol);
        const val = String(cell.value ?? "");

        if (val.includes("Despido")) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: DP_BG },
          };
          cell.font = {
            name: "Calibri",
            size: 11,
            bold: true,
            color: { argb: DP_FONT },
          };
          cell.alignment = { horizontal: "center", vertical: "middle" };
        }
      }
    }

    const widths = [{ width: 10 }, { width: 10 }, { width: 30 }, { width: 16 }];
    days.forEach((_, i) => {
      widths.push({ width: 22 }, { width: 12 }, { width: 12 }, { width: 10 }, { width: 10 });
      for (let j = 1; j <= maxPermsByDay[i]; j++) {
        widths.push({ width: 12 }, { width: 12 });
      }
      if (includeDispatchByDay[i]) widths.push({ width: 12 });
    });
    widths.push({ width: 60 });
    ws.columns = widths;

    const buf = await workbook.xlsx.writeBuffer();
    const filename = `asistencia_semanal_${isActive == 0 ? "inactivos" : "activos"}_semana${selectedWeek}_${year}.xlsx`;
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.send(buf);
  } catch (err) {
    console.error("Error al exportar asistencia semanal:", err.stack || err.message);
    res.status(500).send({
      message: `Error interno del servidor al generar el archivo Excel: ${err.message}`,
    });
  }
};

// Exportar asistencia por rango de fechas (Fechas en columnas, Empleados en filas)
exports.exportDateRangeAttendance = async (req, res) => {
  try {
    const { startDate, endDate, isActive } = req.body;

    // ===== VALIDACIÓN =====
    if (!startDate || typeof startDate !== "string") {
      return res.status(400).json({
        message: "startDate is required and must be a string (YYYY-MM-DD format)",
      });
    }
    if (!endDate || typeof endDate !== "string") {
      return res.status(400).json({
        message: "endDate is required and must be a string (YYYY-MM-DD format)",
      });
    }

    const start = dayjs(startDate, "YYYY-MM-DD", true);
    const end = dayjs(endDate, "YYYY-MM-DD", true);

    if (!start.isValid()) {
      return res.status(400).json({
        message: `startDate is invalid. Expected format: YYYY-MM-DD`,
      });
    }
    if (!end.isValid()) {
      return res.status(400).json({
        message: `endDate is invalid. Expected format: YYYY-MM-DD`,
      });
    }
    if (end.isBefore(start)) {
      return res.status(400).json({
        message: "endDate must be greater than or equal to startDate",
      });
    }

    const activeStatus = isActive !== undefined ? Number(isActive) : 1;

    // ===== GENERAR ARRAY DE FECHAS =====
    const dates = [];
    let current = start;
    while (current.isSameOrBefore(end)) {
      dates.push(current.format("YYYY-MM-DD"));
      current = current.add(1, "day");
    }

    // ===== QUERIES A BD =====
    // 1. Empleados
    const [employees] = await db.query(
      `SELECT 
        employeeID,
        TRIM(CONCAT(firstName, ' ', COALESCE(middleName, ''), ' ', lastName)) AS employeeName
      FROM employees_emp
      WHERE isActive = ?
      ORDER BY employeeID`,
      [activeStatus]
    );

    // 2. Asistencia del rango
    const [attendanceData] = await db.query(
      `SELECT 
        employeeID,
        DATE(date) AS attendanceDate,
        TIME_FORMAT(entryTime, '%H:%i:%s') AS entryTime,
        TIME_FORMAT(exitTime, '%H:%i:%s') AS exitTime
      FROM h_attendance_emp
      WHERE DATE(date) >= ? AND DATE(date) <= ?
      ORDER BY employeeID, date`,
      [startDate, endDate]
    );

    // 3. Permisos aprobados del rango
    const [permissions] = await db.query(
      `SELECT 
        employeeID,
        DATE(date) AS permissionDate,
        request,
        isApproved
      FROM permissionattendance_emp
      WHERE DATE(date) >= ? AND DATE(date) <= ?
        AND isApproved = 1
      ORDER BY employeeID, date`,
      [startDate, endDate]
    );

    // ===== MAPAS PARA O(1) LOOKUP =====
    // attendanceMap: key = "employeeID_date"
    const attendanceMap = new Map();
    for (const record of attendanceData) {
      const key = `${record.employeeID}_${record.attendanceDate}`;
      attendanceMap.set(key, record);
    }

    // permissionMap: key = "employeeID_date", value = array de permisos
    const permissionMap = new Map();
    for (const record of permissions) {
      const key = `${record.employeeID}_${record.permissionDate}`;
      if (!permissionMap.has(key)) {
        permissionMap.set(key, []);
      }
      permissionMap.get(key).push(record);
    }

    // ===== FUNCIÓN PARA DETERMINAR STATUS =====
    const getStatus = (employeeID, date) => {
      const attKey = `${employeeID}_${date}`;
      const permKey = `${employeeID}_${date}`;

      const attendance = attendanceMap.get(attKey);
      const permList = permissionMap.get(permKey) || [];

      // Prioridad: Presente > Permiso > Ausente
      if (attendance && attendance.entryTime) {
        return "Presente";
      }

      if (permList.length > 0 && !attendance) {
        // Si tiene permiso aprobado pero no marcó
        const permTypes = permList
          .map((p) => (p.request === 0 ? "Diferido" : "Solicitado"))
          .join("/");
        return `Permiso (${permTypes})`;
      }

      if (!attendance) {
        return "Ausente";
      }

      return "Sin datos";
    };

    // ===== CREAR EXCEL =====
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("Asistencia por Fecha");

    // Colores
    const HEADER_BG = "4472C4";
    const HEADER_TEXT = "FFFFFF";
    const GREEN = "C6EFCE"; // Presente
    const RED = "FFC7CE"; // Ausente
    const YELLOW = "FFEB9C"; // Permiso
    const GRAY = "E7E6E6"; // Sin datos

    // ===== HEADER =====
    const headerRow = ["ID", "Empleado"];
    for (const date of dates) {
      const dayName = dayjs(date).locale("es").format("ddd"); // lun, mar, mié, jue, vie, sab, dom
      headerRow.push(`${date}\n${dayName}`);
    }

    ws.addRow(headerRow);

    // Aplicar estilos al header
    const headerRowNum = ws.rowCount;
    for (let col = 1; col <= headerRow.length; col++) {
      const cell = ws.getCell(headerRowNum, col);
      cell.font = { bold: true, color: { argb: HEADER_TEXT }, size: 11 };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: HEADER_BG },
      };
      cell.alignment = { horizontal: "center", vertical: "center", wrapText: true };
      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
      };
    }

    // ===== DATOS =====
    for (const employee of employees) {
      const row = [employee.employeeID, employee.employeeName];

      for (const date of dates) {
        const status = getStatus(employee.employeeID, date);
        row.push(status);
      }

      const dataRowNum = ws.addRow(row).number;

      // Aplicar estilos a datos
      for (let col = 1; col <= row.length; col++) {
        const cell = ws.getCell(dataRowNum, col);
        const value = cell.value;

        cell.alignment = { horizontal: "center", vertical: "center" };
        cell.border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" },
        };
        cell.font = { size: 10 };

        // Color según status
        if (value === "Presente") {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: GREEN },
          };
        } else if (value === "Ausente") {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: RED },
          };
        } else if (value && value.toString().includes("Permiso")) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: YELLOW },
          };
        } else if (value === "Sin datos") {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: GRAY },
          };
        }
      }
    }

    // ===== ANCHO DE COLUMNAS Y FROZEN PANES =====
    ws.getColumn(1).width = 10; // ID
    ws.getColumn(2).width = 30; // Empleado
    for (let i = 3; i <= headerRow.length; i++) {
      ws.getColumn(i).width = 16; // Fechas
    }

    // Congelar panel: 2 columnas (ID + Empleado) y 1 fila (header)
    ws.views = [
      {
        xSplit: 2,
        ySplit: 1,
        topLeftCell: "C2",
        activeCell: "C2",
        state: "frozen",
      },
    ];

    // ===== ENVIAR ARCHIVO =====
    const buffer = await workbook.xlsx.writeBuffer();
    const filename = `asistencia_${startDate}_a_${endDate}_${
      activeStatus === 1 ? "activos" : "inactivos"
    }.xlsx`;

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${filename}"`
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.send(buffer);
  } catch (err) {
    console.error(
      "Error al exportar asistencia por rango de fechas:",
      err.stack || err.message
    );
    res.status(500).json({
      message: `Error interno del servidor: ${err.message}`,
    });
  }
};











