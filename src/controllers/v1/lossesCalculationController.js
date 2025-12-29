const LossesCalculationData = require('../../models/v1/lossesCalculation.model');
const moment = require('moment');
const MainClient = require('../../models/v1/mainClient.model');
const SubClient = require('../../models/v1/subClient.model');
const MeterData = require('../../models/v1/meterData.model');
const PartClient = require('../../models/v1/partClient.model');
const logger = require('../../utils/logger');  // Importing the logger
const ExcelJS = require('exceljs');


/**
 * Request body (new format):
 * {
 *   mainClientId: "....",
 *   month: 1,
 *   year: 2025,
 *   // optional legacy fallback:
 *   SLDCGROSSINJECTION: 356.889,
 *   SLDCGROSSDRAWL: -3.333,
 *   // MAIN ENTRY (preferred)
 *   mainEntry: {
 *     approvedInjection: 356.889, // B9
 *     approvedDrawl: -3.333,      // L9 (signed negative or positive? we handle as number)
 *     discom: {                   // C9..I9
 *       DGVCL: 356.883,
 *       MGVCL: 0,
 *       PGVCL: 0,
 *       UGVCL: 0,
 *       TAECO: 0,
 *       TSECO: 0,
 *       TEL: 0
 *     }
 *   }
 * }
 */
// ====================== Losses Calculation Controller (TS/JS) ======================

// ====================== Losses Calculation Controller (TypeScript) ======================

// ---------- Helpers ----------
function pluckNumber(v, def = null) {
  if (v == null) return def;
  if (typeof v === "number") return Number.isFinite(v) ? v : def;
  const n = parseFloat(String(v).replace(/,/g, ""));
  return Number.isFinite(n) ? n : def;
}

function round3(x) {
  if (!Number.isFinite(x)) return 0;
  return Number((Math.round((x + Number.EPSILON) * 1000) / 1000).toFixed(3));
}

function getParam(obj, ...keys) {
  const src = obj || {};
  for (const k of keys) if (src[k] != null) return src[k];

  const lower = Object.create(null);
  for (const k of Object.keys(src)) lower[k.toLowerCase()] = k;

  for (const k of keys) {
    const hit = lower[String(k).toLowerCase()];
    if (hit && src[hit] != null) return src[hit];
  }
  return undefined;
}

function getActiveEnergy(entry) {
  const p = entry?.parameters ?? {};
  const candidates = [
    "Bidirectional Active(I-E)",
    "Bidirectional Active (I-E)",
    "I-E",
    "I - E",
    "Net Active",
    "NET ACTIVE",
    "Net(ACTIVE)",
    "Net (Active)",
  ];
  let val = getParam(p, ...candidates);
  if (val == null) {
    const k = Object.keys(p).find(
      (k) => /active/i.test(k) && /(net|bidirectional|i-?e)/i.test(k)
    );
    if (k) val = p[k];
  }
  return pluckNumber(val);
}

function getDateStr(entry) {
  return String(
    getParam(entry?.parameters ?? {}, "Date", "DATE", "date") ?? ""
  );
}

function getTimeStr(entry) {
  return String(
    getParam(
      entry?.parameters ?? {},
      "Interval Start",
      "IntervalStart",
      "Start Time",
      "Time",
      "START TIME"
    ) ?? ""
  );
}

// ---------- Controller ----------

/**
 * Request body (new format):
 * {
 *   mainClientId: "....",
 *   month: 1,
 *   year: 2025,
 *   // optional legacy fallback:
 *   SLDCGROSSINJECTION: 356.889,
 *   SLDCGROSSDRAWL: -3.333,
 *   // MAIN ENTRY (preferred)
 *   mainEntry: {
 *     approvedInjection: 356.889, // B9
 *     approvedDrawl: -3.333,      // L9 (signed negative or positive? we handle as number)
 *     discom: {                   // C9..I9
 *       DGVCL: 356.883,
 *       MGVCL: 0,
 *       PGVCL: 0,
 *       UGVCL: 0,
 *       TAECO: 0,
 *       TSECO: 0,
 *       TEL: 0
 *     }
 *   }
 * }
 */
// ====================== Losses Calculation Controller (TS/JS) ======================

// ====================== Losses Calculation Controller (TypeScript) ======================

// ---------- Helpers ----------
function pluckNumber(v, def = null) {
  if (v == null) return def;
  if (typeof v === "number") return Number.isFinite(v) ? v : def;
  const n = parseFloat(String(v).replace(/,/g, ""));
  return Number.isFinite(n) ? n : def;
}

function round3(x) {
  if (!Number.isFinite(x)) return 0;
  return Number((Math.round((x + Number.EPSILON) * 1000) / 1000).toFixed(3));
}

function getParam(obj, ...keys) {
  const src = obj || {};
  for (const k of keys) if (src[k] != null) return src[k];

  const lower = Object.create(null);
  for (const k of Object.keys(src)) lower[k.toLowerCase()] = k;

  for (const k of keys) {
    const hit = lower[String(k).toLowerCase()];
    if (hit && src[hit] != null) return src[hit];
  }
  return undefined;
}

function getActiveEnergy(entry) {
  const p = entry?.parameters ?? {};
  const candidates = [
    "Bidirectional Active(I-E)",
    "Bidirectional Active (I-E)",
    "I-E",
    "I - E",
    "Net Active",
    "NET ACTIVE",
    "Net(ACTIVE)",
    "Net (Active)",
  ];
  let val = getParam(p, ...candidates);
  if (val == null) {
    const k = Object.keys(p).find(
      (k) => /active/i.test(k) && /(net|bidirectional|i-?e)/i.test(k)
    );
    if (k) val = p[k];
  }
  return pluckNumber(val);
}

function getDateStr(entry) {
  return String(
    getParam(entry?.parameters ?? {}, "Date", "DATE", "date") ?? ""
  );
}

function getTimeStr(entry) {
  return String(
    getParam(
      entry?.parameters ?? {},
      "Interval Start",
      "IntervalStart",
      "Start Time",
      "Time",
      "START TIME"
    ) ?? ""
  );
}

// ---------- Controller ----------
const generateLossesCalculation = async (req, res) => {
  try {
    const {
      mainClientId,
      month,
      year,
      SLDCGROSSINJECTION,
      SLDCGROSSDRAWL,
      mainEntry, // { approvedInjection, approvedDrawl, discom: {...} }
    } = req.body;

    if (!mainClientId || !month || !year) {
      return res.status(400).json({
        message: "Missing required parameters: mainClientId, month, year.",
      });
    }

    // MAIN ENTRY / SLDC values
    const approvedInjection =
      mainEntry?.approvedInjection != null
        ? pluckNumber(mainEntry.approvedInjection, null)
        : SLDCGROSSINJECTION != null
          ? pluckNumber(SLDCGROSSINJECTION, null)
          : null;

    const approvedDrawl =
      mainEntry?.approvedDrawl != null
        ? pluckNumber(mainEntry.approvedDrawl, null)
        : SLDCGROSSDRAWL != null
          ? pluckNumber(SLDCGROSSDRAWL, null)
          : null;

    const discomTargetsRaw = mainEntry?.discom || {};
    const discomKeys = ["DGVCL", "MGVCL", "PGVCL", "UGVCL", "TAECO", "TSECO", "TEL"];
    const discomTargets = Object.fromEntries(
      discomKeys.map((k) => [k, pluckNumber(discomTargetsRaw[k], 0)])
    );

    // Existing doc reuse (if SLDC values match)
    const existingCalculation = await LossesCalculationData.findOne({
      mainClientId,
      month,
      year,
      ...(approvedInjection !== null && {
        SLDCGROSSINJECTION: approvedInjection,
      }),
      ...(approvedDrawl !== null && { SLDCGROSSDRAWL: approvedDrawl }),
    });
    if (existingCalculation) {
      existingCalculation.updatedAt = new Date();
      await existingCalculation.save();
      return res.status(200).json({
        message: "Existing calculation data retrieved successfully.",
        data: existingCalculation,
      });
    }

    // Entities
    const mainClientData = await MainClient.findById(mainClientId);
    if (!mainClientData) {
      return res.status(404).json({ message: "Main Client not found" });
    }

    const subClients = await SubClient.find({ mainClient: mainClientId });
    if (!subClients.length) {
      return res.status(404).json({ message: "No Sub Clients found" });
    }

    const partClientsData = {};
    await Promise.all(
      subClients.map(async (s) => {
        try {
          partClientsData[s._id] = await PartClient.find({ subClient: s._id });
        } catch (e) {
          logger?.error?.(
            `Error loading part-clients for ${s.name}: ${e.message}`
          );
          partClientsData[s._id] = [];
        }
      })
    );

    // Meter data
    const clientsUsingCheckMeter = [];
    const subClientsUsingCheckMeter = [];

    let mainClientMeterData = await MeterData.find({
      meterNumber: mainClientData.abtMainMeter?.meterNumber,
      month,
      year,
    });
    if (!mainClientMeterData.length && mainClientData.abtCheckMeter?.meterNumber) {
      mainClientMeterData = await MeterData.find({
        meterNumber: mainClientData.abtCheckMeter.meterNumber,
        month,
        year,
      });
      if (mainClientMeterData.length) clientsUsingCheckMeter.push(mainClientData.name);
    }
    if (!mainClientMeterData.length) {
      return res.status(400).json({
        message:
          "Meter data missing for Main Client. Both abtMainMeter and abtCheckMeter files are missing. Calculation cannot proceed.",
      });
    }

    const subClientMeterData = {};
    const missingSubClientMeters = [];
    await Promise.all(
      subClients.map(async (s) => {
        let data = await MeterData.find({
          meterNumber: s.abtMainMeter?.meterNumber,
          month,
          year,
        });
        if (!data.length && s.abtCheckMeter?.meterNumber) {
          data = await MeterData.find({
            meterNumber: s.abtCheckMeter.meterNumber,
            month,
            year,
          });
          if (data.length) subClientsUsingCheckMeter.push(s.name);
        }
        if (!data.length) missingSubClientMeters.push(s.name);
        subClientMeterData[s._id] = data;
      })
    );
    if (missingSubClientMeters.length) {
      return res.status(400).json({
        message: `Meter data missing for Sub Clients: ${missingSubClientMeters.join(
          ", "
        )}. Calculation cannot proceed.`,
      });
    }

    // Init doc
    const doc = new LossesCalculationData({
      mainClientId,
      month,
      year,
      SLDCGROSSINJECTION: approvedInjection ?? undefined,
      SLDCGROSSDRAWL: approvedDrawl ?? undefined,

      DGVCL: Number.isFinite(discomTargets.DGVCL) ? discomTargets.DGVCL : undefined,
      MGVCL: Number.isFinite(discomTargets.MGVCL) ? discomTargets.MGVCL : undefined,
      PGVCL: Number.isFinite(discomTargets.PGVCL) ? discomTargets.PGVCL : undefined,
      UGVCL: Number.isFinite(discomTargets.UGVCL) ? discomTargets.UGVCL : undefined,
      TAECO: Number.isFinite(discomTargets.TAECO) ? discomTargets.TAECO : undefined,
      TSECO: Number.isFinite(discomTargets.TSECO) ? discomTargets.TSECO : undefined,
      TEL: Number.isFinite(discomTargets.TEL) ? discomTargets.TEL : undefined,

      mainClient: {
        meterNumber: mainClientMeterData[0].meterNumber,
        meterType: mainClientMeterData[0].meterType,
        mainClientDetail: {
          name: mainClientData.name,
          subTitle: mainClientData.subTitle,
          abtMainMeter: mainClientData.abtMainMeter,
          abtCheckMeter: mainClientData.abtCheckMeter,
          voltageLevel: mainClientData.voltageLevel,
          acCapacityKw: mainClientData.acCapacityKw,
          dcCapacityKwp: mainClientData.dcCapacityKwp,
          noOfModules: mainClientData.noOfModules,
          ctptSrNo: mainClientData.ctptSrNo,
          ctRatio: mainClientData.ctRatio,
          ptRatio: mainClientData.ptRatio,
          mf: mainClientData.mf,
          sharingPercentage: mainClientData.sharingPercentage,
          contactNo: mainClientData.contactNo,
          email: mainClientData.email,
        },
        grossInjectionMWH: 0,
        drawlMWH: 0,
        netInjectionMWH: 0,
        mainClientMeterDetails: [],
      },

      subClient: [],
      subClientoverall: { overallGrossInjectedUnits: 0, grossDrawlUnits: 0 },
      difference: { diffInjectedUnits: 0, diffDrawlUnits: 0 },
      audit: {},
    });

    // MAIN raw
    const mainPn = Number.isFinite(mainClientData.pn) ? mainClientData.pn : -1;
    const mainMf = Number.isFinite(mainClientData.mf) ? mainClientData.mf : 1;

    let mainRawPos = 0;
    let mainRawNeg = 0;

    mainClientMeterData.forEach((m) => {
      m.dataEntries.forEach((e) => {
        const aE = getActiveEnergy(e);
        if (!Number.isFinite(aE)) return;
        const vRaw = (aE * mainMf * mainPn) / 1000;
        const date = getDateStr(e);
        const time = getTimeStr(e);

        doc.mainClient.mainClientMeterDetails.push({
          date,
          time,
          grossInjectedUnitsTotal: vRaw,
          helper: { raw: vRaw },
        });

        if (vRaw > 0) mainRawPos += vRaw;
        else mainRawNeg += vRaw;
      });
    });
    doc.audit.mainRaw = { pos: mainRawPos, neg: mainRawNeg };

    // SUBS raw + helper.raw
    for (const s of subClients) {
      const data = subClientMeterData[s._id];
      const sPn = Number.isFinite(s.pn) ? s.pn : -1;
      const sMf = Number.isFinite(s.mf) ? s.mf : 1;

      const subBlock = {
        name: s.name,
        divisionName: s.divisionName,
        consumerNo: s.consumerNo,
        contactNo: s.contactNo,
        email: s.email,
        subClientId: s._id,
        meterNumber: data[0].meterNumber,
        meterType: data[0].meterType,
        discom: s.discom,
        voltageLevel: s.voltageLevel,
        ctptSrNo: s.ctptSrNo,
        ctRatio: s.ctRatio,
        ptRatio: s.ptRatio,
        mf: s.mf,
        acCapacityKw: s.acCapacityKw,
        subClientsData: {
          grossInjectionMWH: 0,
          drawlMWH: 0,
          netInjectionMWH: 0,
          grossInjectionMWHAfterLosses: 0,
          drawlMWHAfterLosses: 0,
          netInjectionMWHAfterLosses: 0,
          weightageGrossInjecting: 0,
          weightageGrossDrawl: 0,
          lossesInjectedUnits: 0,
          inPercentageOfLossesInjectedUnits: 0,
          lossesDrawlUnits: 0,
          inPercentageOfLossesDrawlUnits: 0,
          partclient: [],
          subClientMeterData: [],
        },
      };

      const pcs = partClientsData[s._id] || [];
      if (pcs.length) {
        subBlock.subClientsData.partclient = pcs.map((p) => ({
          divisionName: p.divisionName,
          consumerNo: p.consumerNo,
          sharingPercentage: pluckNumber(p.sharingPercentage, 0),
          grossInjectionMWHAfterLosses: 0,
          drawlMWHAfterLosses: 0,
          netInjectionMWHAfterLosses: 0,
          grossInjectionMWH: 0,
          drawlMWH: 0,
          netInjectionMWH: 0,
          weightageGrossInjecting: 0,
          weightageGrossDrawl: 0,
          lossesInjectedUnits: 0,
          inPercentageOfLossesInjectedUnits: 0,
          lossesDrawlUnits: 0,
          inPercentageOfLossesDrawlUnits: 0,
        }));
      }

      data.forEach((m) => {
        m.dataEntries.forEach((e) => {
          const aE = getActiveEnergy(e);
          if (!Number.isFinite(aE)) return;

          const v = (aE * sMf * sPn) / 1000;
          const date = getDateStr(e);
          const time = getTimeStr(e);

          subBlock.subClientsData.subClientMeterData.push({
            date,
            time,
            grossInjectedUnitsTotal: v,
            netTotalAfterLosses: 0,
            partclient: [],
            helper: { raw: v },
          });

          if (v > 0) subBlock.subClientsData.grossInjectionMWH += v;
          else subBlock.subClientsData.drawlMWH += v;
        });
      });

      subBlock.subClientsData.netInjectionMWH =
        subBlock.subClientsData.grossInjectionMWH +
        subBlock.subClientsData.drawlMWH;

      doc.subClient.push(subBlock);
      doc.subClientoverall.overallGrossInjectedUnits +=
        subBlock.subClientsData.grossInjectionMWH;
      doc.subClientoverall.grossDrawlUnits +=
        subBlock.subClientsData.drawlMWH;
    }

    // MAIN scaling by SLDC (fPos / fNeg)
    let fPos = 1;
    let fNeg = 1;
    if (approvedInjection != null && mainRawPos !== 0)
      fPos = approvedInjection / mainRawPos;
    if (approvedDrawl != null && mainRawNeg !== 0)
      fNeg = approvedDrawl / mainRawNeg;

    doc.audit.mainScale = { fPos, fNeg };

    let mainGrossAdj = 0;
    let mainDrawlAdj = 0;
    doc.mainClient.mainClientMeterDetails =
      doc.mainClient.mainClientMeterDetails.map((row) => {
        const v = row.helper?.raw ?? row.grossInjectedUnitsTotal;
        const adj = v > 0 ? v * fPos : v * fNeg;
        const next = { ...row, grossInjectedUnitsTotal: adj };
        if (adj > 0) mainGrossAdj += adj;
        else mainDrawlAdj += adj;
        return next;
      });
    doc.mainClient.grossInjectionMWH = mainGrossAdj;
    doc.mainClient.drawlMWH = mainDrawlAdj;
    doc.mainClient.netInjectionMWH = mainGrossAdj + mainDrawlAdj;

    // Differences main vs subs (raw on subs)
    doc.difference.diffInjectedUnits =
      doc.subClientoverall.overallGrossInjectedUnits -
      doc.mainClient.grossInjectionMWH;
    doc.difference.diffDrawlUnits =
      doc.subClientoverall.grossDrawlUnits - doc.mainClient.drawlMWH;

    // Monthly sums (raw)
    const overallPos = doc.subClient.reduce(
      (a, s) => a + s.subClientsData.grossInjectionMWH,
      0
    );
    const overallNeg = doc.subClient.reduce(
      (a, s) => a + s.subClientsData.drawlMWH,
      0
    );
    doc.audit.subsPositiveSum = overallPos;
    doc.audit.subsNegativeSum = overallNeg;

    // ----- helper.allocatedGroup: interval allocation per main step -----
    const mainByKey = new Map();
    doc.mainClient.mainClientMeterDetails.forEach((row) => {
      const key = `${row.date}__${row.time}`;
      mainByKey.set(key, row.grossInjectedUnitsTotal);
    });

    const groups = new Map();
    doc.subClient.forEach((sc) => {
      sc.subClientsData.subClientMeterData.forEach((row) => {
        const key = `${row.date}__${row.time}`;
        let g = groups.get(key);
        if (!g) {
          g = { sumRaw: 0, rows: [] };
          groups.set(key, g);
        }
        const raw = row.helper?.raw ?? row.grossInjectedUnitsTotal;
        g.sumRaw += raw;
        g.rows.push(row);
      });
    });

    groups.forEach((g, key) => {
      const s = g.sumRaw;
      const target = mainByKey.get(key);

      if (!Number.isFinite(s) || s === 0 || !Number.isFinite(target)) {
        g.rows.forEach((row) => {
          const raw = row.helper?.raw ?? row.grossInjectedUnitsTotal;
          row.helper = row.helper || {};
          row.helper.allocatedGroup = raw;
          row.helper.discomScaled = raw;
        });
        return;
      }

      let sumAlloc = 0;
      g.rows.forEach((row) => {
        const raw = row.helper?.raw ?? row.grossInjectedUnitsTotal;
        const alloc = raw + (raw / s) * (target - s);
        row.helper = row.helper || {};
        row.helper.allocatedGroup = alloc;
        sumAlloc += alloc;
      });

      const delta = target - sumAlloc;
      if (g.rows.length && Number.isFinite(delta) && Math.abs(delta) > 0) {
        g.rows[0].helper.allocatedGroup += delta;
      }
    });

    // ----- helper.discomScaled: per-sub scaling so sum(discomScaled) == sum(raw) -----
    doc.subClient.forEach((sc) => {
      const rows = sc.subClientsData.subClientMeterData;

      let posE = 0;
      let posF = 0;
      const positiveRows = [];

      rows.forEach((row) => {
        const raw = row.helper?.raw ?? row.grossInjectedUnitsTotal;
        const alloc =
          row.helper?.allocatedGroup ??
          row.helper?.raw ??
          row.grossInjectedUnitsTotal;
        if (raw > 0) {
          posE += raw;
          positiveRows.push(row);
        }
        if (alloc > 0) posF += alloc;
      });

      const scale = posF !== 0 ? posE / posF : 1;

      let sumScaled = 0;
      positiveRows.forEach((row) => {
        const alloc =
          row.helper?.allocatedGroup ??
          row.helper?.raw ??
          row.grossInjectedUnitsTotal;
        const scaled = alloc * scale;
        row.helper.discomScaled = scaled;
        sumScaled += scaled;
      });

      if (positiveRows.length && posE !== 0) {
        const delta = posE - sumScaled;
        positiveRows[0].helper.discomScaled += delta;
      }

      rows.forEach((row) => {
        const raw = row.helper?.raw ?? row.grossInjectedUnitsTotal;
        if (!(raw > 0)) {
          row.helper.discomScaled = raw;
        }
      });
    });

    // Update grossInjectedUnitsTotal to use discomScaled for sub clients
    doc.subClient.forEach((sc) => {
      sc.subClientsData.subClientMeterData.forEach((row) => {
        if (row.helper && row.helper.discomScaled !== undefined) {
          row.grossInjectedUnitsTotal = row.helper.discomScaled;
        }
      });
    });

    // ----- Loss percentages & first pass after-losses -----
    doc.subClient.forEach((sc) => {
      const d = sc.subClientsData;

      d.weightageGrossInjecting = overallPos
        ? (d.grossInjectionMWH / overallPos) * 100
        : 0;
      d.weightageGrossDrawl = overallNeg
        ? (d.drawlMWH / overallNeg) * 100
        : 0;

      d.lossesInjectedUnits =
        doc.difference.diffInjectedUnits * (d.weightageGrossInjecting / 100);
      d.lossesDrawlUnits =
        doc.difference.diffDrawlUnits * (d.weightageGrossDrawl / 100);

      d.inPercentageOfLossesInjectedUnits = d.grossInjectionMWH
        ? (d.lossesInjectedUnits / d.grossInjectionMWH) * 100
        : 0;
      d.inPercentageOfLossesDrawlUnits = d.drawlMWH
        ? (d.lossesDrawlUnits / d.drawlMWH) * 100
        : 0;

      if (Array.isArray(d.partclient) && d.partclient.length) {
        d.partclient.forEach((pc) => {
          const pct = (pc.sharingPercentage || 0) / 100;
          pc.grossInjectionMWH = d.grossInjectionMWH * pct;
          pc.drawlMWH = d.drawlMWH * pct;
          pc.netInjectionMWH = d.netInjectionMWH * pct;

          pc.weightageGrossInjecting = d.weightageGrossInjecting * pct;
          pc.weightageGrossDrawl = d.weightageGrossDrawl * pct;
          pc.lossesInjectedUnits = d.lossesInjectedUnits * pct;
          pc.inPercentageOfLossesInjectedUnits =
            d.inPercentageOfLossesInjectedUnits;
          pc.lossesDrawlUnits = d.lossesDrawlUnits * pct;
          pc.inPercentageOfLossesDrawlUnits =
            d.inPercentageOfLossesDrawlUnits;
        });
      }
    });

    // First pass: per-row net after losses from percentages
    doc.subClient.forEach((sc) => {
      const d = sc.subClientsData;
      let gAfter = 0;
      let dAfter = 0;

      d.subClientMeterData.forEach((row) => {
        const v = row.grossInjectedUnitsTotal;
        const lossPct =
          v > 0
            ? d.inPercentageOfLossesInjectedUnits
            : d.inPercentageOfLossesDrawlUnits;
        const after = ((v * (lossPct / 100)) - v) * -1;
        row.netTotalAfterLosses = after;

        if (after > 0) gAfter += after;
        else dAfter += after;
      });

      d.grossInjectionMWHAfterLosses = gAfter;
      d.drawlMWHAfterLosses = dAfter;
      d.netInjectionMWHAfterLosses = gAfter + dAfter;

      if (Array.isArray(d.partclient) && d.partclient.length) {
        d.partclient.forEach((pc) => {
          const pct = (pc.sharingPercentage || 0) / 100;
          pc.grossInjectionMWHAfterLosses = gAfter * pct;
          pc.drawlMWHAfterLosses = dAfter * pct;
          pc.netInjectionMWHAfterLosses =
            pc.grossInjectionMWHAfterLosses +
            pc.drawlMWHAfterLosses;
        });
      }
    });

    // ------------------------------------------------------------------
    // NEW STEP: Global DISCOM match on net (positive) units
    // Sum of all positive netTotalAfterLosses across subs
    //   = Sum of MAIN ENTRY discom targets (e.g. 356.883)
    // ------------------------------------------------------------------
    const sumDiscomTargetsPos = Object.values(discomTargets).reduce(
      (acc, v) => (Number.isFinite(v) ? acc + v : acc),
      0
    );

    let totalNetPos = 0;
    doc.subClient.forEach((sc) => {
      sc.subClientsData.subClientMeterData.forEach((row) => {
        if (row.netTotalAfterLosses > 0) {
          totalNetPos += row.netTotalAfterLosses;
        }
      });
    });

    if (sumDiscomTargetsPos > 0 && totalNetPos > 0) {
      const scaleNet = sumDiscomTargetsPos / totalNetPos;
      doc.audit = doc.audit || {};
      doc.audit.discomNetScale = {
        sumDiscomTargetsPos,
        totalNetPos,
        scaleNet,
      };

      // Apply scaling & recompute after-loss totals
      doc.subClient.forEach((sc) => {
        const d = sc.subClientsData;
        let gAfter = 0;
        let dAfter = 0;

        d.subClientMeterData.forEach((row) => {
          if (row.netTotalAfterLosses > 0) {
            row.netTotalAfterLosses = row.netTotalAfterLosses * scaleNet;
          }

          if (row.netTotalAfterLosses > 0) gAfter += row.netTotalAfterLosses;
          else dAfter += row.netTotalAfterLosses;
        });

        d.grossInjectionMWHAfterLosses = gAfter;
        d.drawlMWHAfterLosses = dAfter;
        d.netInjectionMWHAfterLosses = gAfter + dAfter;

        // Keep part-clients consistent with final after-loss totals
        if (Array.isArray(d.partclient) && d.partclient.length) {
          d.partclient.forEach((pc) => {
            const pct = (pc.sharingPercentage || 0) / 100;
            pc.grossInjectionMWHAfterLosses = d.grossInjectionMWHAfterLosses * pct;
            pc.drawlMWHAfterLosses = d.drawlMWHAfterLosses * pct;
            pc.netInjectionMWHAfterLosses =
              pc.grossInjectionMWHAfterLosses +
              pc.drawlMWHAfterLosses;
          });
        }
      });
    }

    // SLDC diffs
    if (approvedInjection != null) {
      doc.SLDCGROSSINJECTION = approvedInjection;
      doc.mainClient.asperApprovedbySLDCGROSSINJECTION =
        approvedInjection - doc.mainClient.grossInjectionMWH;
    }
    if (approvedDrawl != null) {
      doc.SLDCGROSSDRAWL = approvedDrawl;
      doc.mainClient.asperApprovedbySLDCGROSSDRAWL =
        approvedDrawl - doc.mainClient.drawlMWH;
    }

    // ------------------------------------------------------------------
    // DISCOM totals & Excess Injection PPA (match MAIN ENTRY sheet)
    // We rely on mainEntry.discom for month totals:
    //   perDiscomTotals[DGVCL] = 356.883, etc.
    //   ExcessInjectionPPA = SLDCGROSSINJECTION - SUM(discomTargets)
    // ------------------------------------------------------------------
    const credited = {};
    discomKeys.forEach((k) => {
      const v = discomTargets[k]; // MAIN ENTRY J9 side
      if (Number.isFinite(v)) {
        credited[k] = round3(v);
      }
    });
    doc.perDiscomTotals = credited;

    const totalCredited = Object.values(credited).reduce(
      (a, b) => a + b,
      0
    );
    const sldcApprovedInj =
      approvedInjection != null ? approvedInjection : totalCredited;

    doc.excessInjectionPPA = round3(sldcApprovedInj - totalCredited);

    // Drawl side (just store energyDrawnFromDiscom)
    doc.energyDrawnFromDiscom =
      approvedDrawl != null
        ? pluckNumber(approvedDrawl, doc.mainClient.drawlMWH)
        : doc.mainClient.drawlMWH;

    // Check-meter info
    const allClientsUsingCheckMeter = [
      ...clientsUsingCheckMeter,
      ...subClientsUsingCheckMeter,
    ];
    if (allClientsUsingCheckMeter.length) {
      doc.clientsUsingCheckMeter = allClientsUsingCheckMeter;
    }

    await doc.save();

    return res.status(200).json({
      message: "Losses Calculation successfully completed.",
      data: doc,
      ...(allClientsUsingCheckMeter.length && {
        clientsUsingCheckMeter: allClientsUsingCheckMeter,
      }),
    });
  } catch (error) {
    logger?.error?.(
      `Error generating Losses Calculation Data: ${error.message}`
    );
    return res.status(500).json({
      message: "Error generating Losses Calculation Data",
      error: error.message,
    });
  }
};



// Get latest 10 losses calculation reports
const getLatestLossesReports = async (req, res) => {
  try {
    // Get latest 10 reports sorted by updatedAt (newest first)
    logger.info("Fetching the latest 10 losses calculation reports, sorted by last update.");

    const latestReports = await LossesCalculationData.find({})
      .sort({ updatedAt: -1 })  // Sorting by updatedAt to get the latest reports
      .limit(10)
      .lean();

    if (!latestReports || latestReports.length === 0) {
      logger.warn("No losses calculation reports found.");
      return res.status(404).json({
        message: 'No losses calculation reports found'
      });
    }

    // Transform the data to include only necessary fields
    const simplifiedReports = latestReports.map(report => {
      return {
        id: report._id,
        month: report.month,
        year: report.year,
        clientName: report.mainClient?.mainClientDetail?.name || 'N/A', // Using 'N/A' if client name is missing
        lastUpdated: report.updatedAt, // The last updated timestamp
        generatedAt: report.createdAt, // Original creation timestamp
        grossInjection: report.mainClient?.grossInjectionMWH || 0,  // Default to 0 if no data available
        totalDrawl: report.mainClient?.drawlMWH || 0,  // Default to 0 if no data available
        reportType: 'Losses'  // Hardcoded as 'Losses' for now
      };
    });

    // Sending the successful response with the processed data
    logger.info("Latest 10 losses calculation reports retrieved successfully.");
    res.status(200).json({
      message: 'Latest 10 losses calculation reports retrieved successfully (sorted by last update)',
      data: simplifiedReports
    });

  } catch (error) {
    // Enhanced error handling with logger for easier troubleshooting
    logger.error(`Error fetching losses reports: ${error.message}`, { error: error.stack });
    res.status(500).json({
      message: 'Error fetching losses calculation reports',
      error: error.message  // Providing the error message to frontend
    });
  }
};


const exportLossesCalculationToExcel = async (lossesCalculationData) => {
  // Create a new workbook
  const workbook = new ExcelJS.Workbook();

  const worksheetSetup = {
    margins: { left: 0.2, right: 0.2, top: 0.2, bottom: 0.2, header: 0.2, footer: 0.2 },
    horizontalCentered: true,
    verticalCentered: false,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    paperSize: 9, // A4
  };

  // Create SUMMARY sheet
  const summarySheet = workbook.addWorksheet("SUMMARY");
  summarySheet.pageSetup = worksheetSetup;

  // Tab color
  summarySheet.properties.tabColor = { argb: "FFFF00" };

  const month =
    lossesCalculationData.month < 10
      ? `0${lossesCalculationData.month}`
      : lossesCalculationData.month;

  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const monthName = monthNames[lossesCalculationData.month - 1];

  // Prepare data rows
  const subClients = lossesCalculationData.subClient;

  // Date helpers
  const monthStr =
    lossesCalculationData.month < 10
      ? `0${lossesCalculationData.month}`
      : lossesCalculationData.month.toString();

  const lastDay = new Date(
    lossesCalculationData.year,
    lossesCalculationData.month,
    0
  ).getDate();

  // Column widths - dynamic based on number of subclients
  // Minimum K column (5 subclients) tak columns, but agar 6+ subclients hain to unke columns add karo
  const numSubClientsForCols = lossesCalculationData.subClient.length;
  const baseColumns = [
    { width: 7 },   // A - Sr. No.
    { width: 30 },  // B - HT Consumer Name
    { width: 13 },  // C - HT Consumer No.
    { width: 22 },  // D - Wheeling Division Office/Location
    { width: 22 },  // E - Wheeling Discom (Main Client)
  ];
  // Add columns for subclients (F onwards) - actual number of subclients
  const summarySubClientColumns = Array.from({ length: numSubClientsForCols }, () => ({ width: 22 }));
  // Add additional gray columns to reach K column (only if <= 5 subclients)
  // Formula: (5 - numSubClients) + 1 = number of gray columns needed to reach K
  // Example: 3 subclients need 3 gray columns (I, J, K), 5 subclients need 1 gray column (K)
  const numGrayColumns = numSubClientsForCols <= 5 ? ((5 - numSubClientsForCols) + 1) : 0;
  const additionalColumns = Array.from({ length: numGrayColumns }, () => ({ width: 22 }));
  summarySheet.columns = [...baseColumns, ...summarySubClientColumns, ...additionalColumns];

  // Helper for fixed 3-decimal text
  const displayExactValue = (value) => {
    if (value === undefined || value === null || isNaN(value)) return "0.000";
    const numValue = Number(value);
    return numValue.toLocaleString("en-US", {
      minimumFractionDigits: 3,
      maximumFractionDigits: 3,
      useGrouping: false,
    });
  };

  // Row 1 spacer
  summarySheet.getRow(1).height = 15;

  // Calculate last column for merged cells - minimum K, but extend if subclients > 5
  const numSubClientsForMerged = lossesCalculationData.subClient.length;
  // Subclients start at F (70), so last subclient column = 70 + numSubClients - 1
  // Minimum K (75), but if > 5 subclients, extend to last subclient column
  const lastMergedColumn = numSubClientsForMerged <= 5 ? 75 : (70 + numSubClientsForMerged - 1);
  const lastMergedColumnChar = String.fromCharCode(lastMergedColumn);

  // Row 2: Titles
  const titleRow = summarySheet.getRow(2);
  titleRow.height = 42;

  summarySheet.mergeCells("A2:B2");
  const summaryTitleCell = summarySheet.getCell("A2");
  summaryTitleCell.value = "SUMMARY SHEET";
  summaryTitleCell.font = { bold: true, size: 16, name: "Times New Roman", color: { argb: "000000" } };
  summaryTitleCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  summaryTitleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "b4c6e7" } };
  summaryTitleCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  summarySheet.mergeCells(`C2:${lastMergedColumnChar}2`);
  const companyCellSummary = summarySheet.getCell("C2");
  const acCapacityMwSummary = (
    (lossesCalculationData.mainClient.mainClientDetail.acCapacityKw || 0) / 1000
  ).toFixed(2);
  companyCellSummary.value = `${lossesCalculationData.mainClient.mainClientDetail.name.toUpperCase()} - ${acCapacityMwSummary} MW AC Generation Details`;
  companyCellSummary.font = { bold: true, size: 16, name: "Times New Roman", color: { argb: "000000" } };
  companyCellSummary.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  companyCellSummary.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "b4c6e7" } };
  companyCellSummary.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  // Row 3: Month
  const monthRow = summarySheet.getRow(3);
  monthRow.height = 30;

  summarySheet.mergeCells("A3:E3");
  const monthLabelCell = summarySheet.getCell("A3");
  monthLabelCell.value = "Month";
  monthLabelCell.font = { bold: true, size: 14, name: "Times New Roman", color: { argb: "FF0000" } };
  monthLabelCell.alignment = { horizontal: "center", vertical: "middle" };
  monthLabelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "b4c6e7" } };
  monthLabelCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  summarySheet.mergeCells(`F3:${lastMergedColumnChar}3`);
  const monthValueCell = summarySheet.getCell("F3");
  monthValueCell.value = `${monthName}-${lossesCalculationData.year.toString().slice(-2)}`;
  monthValueCell.font = { bold: true, size: 18, name: "Times New Roman", color: { argb: "FF0000" } };
  monthValueCell.alignment = { horizontal: "center", vertical: "middle" };
  monthValueCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "b4c6e7" } };
  monthValueCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  // Row 4: Generation Period
  const periodRow = summarySheet.getRow(4);
  periodRow.height = 30;

  summarySheet.mergeCells("A4:E4");
  const periodLabelCell = summarySheet.getCell("A4");
  periodLabelCell.value = "Generation Period";
  periodLabelCell.font = { bold: true, size: 14, name: "Times New Roman", color: { argb: "FF0000" } };
  periodLabelCell.alignment = { horizontal: "center", vertical: "middle" };
  periodLabelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "b4c6e7" } };
  periodLabelCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  const periodEndColumn = lastMergedColumnChar;
  summarySheet.mergeCells(`F4:${periodEndColumn}4`);
  const periodValueCell = summarySheet.getCell("F4");
  const startDate = `01-${monthStr}-${lossesCalculationData.year}`;
  const endDate = `${lastDay}-${monthStr}-${lossesCalculationData.year}`;
  periodValueCell.value = `${startDate} to ${endDate}`;
  periodValueCell.font = { bold: true, size: 18, name: "Times New Roman", color: { argb: "FF0000" } };
  periodValueCell.alignment = { horizontal: "center", vertical: "middle" };
  periodValueCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "b4c6e7" } };
  periodValueCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  // Row 5: CPP header line
  const cppRow = summarySheet.getRow(5);
  cppRow.height = 40;
  summarySheet.mergeCells(`A5:${lastMergedColumnChar}5`);
  const cppCell = summarySheet.getCell("A5");
  cppCell.value = `CPP CLIENTS - ${lossesCalculationData.mainClient.mainClientDetail.name.toUpperCase()} (Lead generator) SOLAR PLANT WITH INJECTION TO ${lossesCalculationData.mainClient.mainClientDetail.subTitle} AT 11kv, ABT METER: ${lossesCalculationData.mainClient.meterNumber}`;
  cppCell.font = { bold: true, size: 12, name: "Times New Roman", color: { argb: "0000cc" } };
  cppCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  cppCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "b4c6e7" } };
  cppCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  // Row 6: SLDC APPROVED strip
  const sldcRow = summarySheet.getRow(6);
  sldcRow.height = 70;

  summarySheet.mergeCells("A6:B6");
  const sldcLabelCell = summarySheet.getCell("A6");
  sldcLabelCell.value = "SLDC APPROVED";
  sldcLabelCell.font = { bold: true, size: 10, name: "Times New Roman" };
  sldcLabelCell.alignment = { horizontal: "center", vertical: "middle" };
  sldcLabelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "92D050" } };
  sldcLabelCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  const totalLabelCell = summarySheet.getCell("C6");
  totalLabelCell.value = "Total";
  totalLabelCell.font = { bold: true, size: 10, name: "Times New Roman" };
  totalLabelCell.alignment = { horizontal: "center", vertical: "middle" };
  totalLabelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "92D050" } };
  totalLabelCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  const feederLabelCell = summarySheet.getCell("D6");
  feederLabelCell.value = "Feeder Name =>";
  feederLabelCell.font = { bold: true, size: 10, name: "Times New Roman" };
  feederLabelCell.alignment = { horizontal: "right", vertical: "middle", wrapText: true };
  feederLabelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
  feederLabelCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  const mainClientCell = summarySheet.getCell("E6");
  mainClientCell.value = `(Lead Generator)\n${lossesCalculationData.mainClient.mainClientDetail.name.toUpperCase()}`;
  mainClientCell.font = { bold: true, size: 10, name: "Times New Roman" };
  mainClientCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  mainClientCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
  mainClientCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  const numSubClients = lossesCalculationData.subClient.length;
  for (let i = 0; i < numSubClients; i++) {
    const col = String.fromCharCode(70 + i); // F onwards
    const cellRef = summarySheet.getCell(`${col}6`);
    cellRef.value = lossesCalculationData.subClient[i]
      ? lossesCalculationData.subClient[i].name.toUpperCase()
      : "";
    cellRef.font = { bold: true, size: 10, name: "Times New Roman" };
    cellRef.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cellRef.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
    cellRef.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
  }

  // Fill intermediate columns with gray background up to K column (if <= 5 subclients)
  // Minimum K column (75) tak gray background, but agar 6+ subclients hain to extra gray column nahi dikhana
  if (numSubClients <= 5) {
    // Agar 5 ya kam subclients hain, to last subclient se K column tak sab columns me gray background dikhao
    const lastSubClientCol = 70 + numSubClients; // Column after last subclient
    const kColumn = 75; // K column
    // Fill all columns from last subclient to K with gray
    for (let colCode = lastSubClientCol; colCode <= kColumn; colCode++) {
      const col = String.fromCharCode(colCode);
      const cellRef = summarySheet.getCell(`${col}6`);
      cellRef.value = "";
      cellRef.font = { bold: true, size: 10, name: "Times New Roman" };
      cellRef.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      cellRef.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
      cellRef.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    }
  }
  // Agar 6+ subclients hain, to extra gray column nahi dikhana

  // Totals (exact values)
  const mainClientGrossInjection = lossesCalculationData.mainClient.grossInjectionMWH || 0;
  const mainClientDrawl = lossesCalculationData.mainClient.drawlMWH || 0;
  const mainClientNetInjection = mainClientGrossInjection + mainClientDrawl;

  const subClientGrossInjection = lossesCalculationData.subClient.reduce(
    (sum, sc) => sum + (sc.subClientsData?.grossInjectionMWH || 0),
    0
  );
  const subClientDrawl = lossesCalculationData.subClient.reduce(
    (sum, sc) => sum + (sc.subClientsData?.drawlMWH || 0),
    0
  );
  const subClientNetInjection = subClientGrossInjection + subClientDrawl;

  // Prefer SLDC if present
  const grossInjectedValue =
    lossesCalculationData.SLDCGROSSINJECTION !== undefined
      ? lossesCalculationData.SLDCGROSSINJECTION
      : mainClientGrossInjection;

  const grossDrawlValue =
    lossesCalculationData.SLDCGROSSDRAWL !== undefined
      ? lossesCalculationData.SLDCGROSSDRAWL
      : mainClientDrawl;

  const netInjectedValue = grossInjectedValue + grossDrawlValue;

  // Calculate totals (same as gray totals row) for green box
  let totalGrossInjectedForGreenBox = 0;
  let totalGrossDrawlForGreenBox = 0;
  lossesCalculationData.subClient.forEach((subClient) => {
    const sc = subClient.subClientsData || {};
    if (sc.partclient && sc.partclient.length > 0) {
      sc.partclient.forEach((pc) => {
        totalGrossInjectedForGreenBox += pc.grossInjectionMWHAfterLosses || 0;
        totalGrossDrawlForGreenBox += pc.drawlMWHAfterLosses || 0;
      });
    } else {
      totalGrossInjectedForGreenBox += sc.grossInjectionMWHAfterLosses || 0;
      totalGrossDrawlForGreenBox += sc.drawlMWHAfterLosses || 0;
    }
  });
  const totalNetInjectedForGreenBox = totalGrossInjectedForGreenBox + totalGrossDrawlForGreenBox;

  // Rows 7–12 (merged A/B/C blocks)
  [7, 8, 9, 10, 11, 12].forEach((r) => (summarySheet.getRow(r).height = 45));
  const greenFill = { type: "pattern", pattern: "solid", fgColor: { argb: "92D050" } };
  const labelFont = { bold: true, size: 10, name: "Times New Roman" };
  const valueFont = { bold: true, size: 10, name: "Times New Roman" };
  const centerMid = { horizontal: "center", vertical: "middle" };
  const leftMidWrap = { horizontal: "left", vertical: "middle", wrapText: true };
  const thinBorder = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  // A7:B8 / C7:C8 - Gross Injected
  summarySheet.mergeCells("A7:B8");
  summarySheet.mergeCells("C7:C8");
  const a7 = summarySheet.getCell("A7");
  a7.value = `Gross Injected Units to ${lossesCalculationData.mainClient.mainClientDetail.subTitle}`;
  a7.font = labelFont; a7.alignment = leftMidWrap; a7.fill = greenFill; a7.border = thinBorder;
  const c7 = summarySheet.getCell("C7");
  c7.value = displayExactValue(totalGrossInjectedForGreenBox);
  c7.font = valueFont; c7.alignment = centerMid; c7.fill = greenFill; c7.border = thinBorder;

  // A9:B10 / C9:C10 - Gross Drawl
  summarySheet.mergeCells("A9:B10");
  summarySheet.mergeCells("C9:C10");
  const a9 = summarySheet.getCell("A9");
  a9.value = `Gross Drawl Units from ${lossesCalculationData.mainClient.mainClientDetail.subTitle}`;
  a9.font = labelFont; a9.alignment = leftMidWrap; a9.fill = greenFill; a9.border = thinBorder;
  const c9 = summarySheet.getCell("C9");
  c9.value = displayExactValue(totalGrossDrawlForGreenBox);
  c9.font = valueFont; c9.alignment = centerMid; c9.fill = greenFill; c9.border = thinBorder;

  // A11:B12 / C11:C12 - Net Injected
  summarySheet.mergeCells("A11:B12");
  summarySheet.mergeCells("C11:C12");
  const a11 = summarySheet.getCell("A11");
  a11.value = `Net Injected Units to ${lossesCalculationData.mainClient.mainClientDetail.subTitle}`;
  a11.font = labelFont; a11.alignment = leftMidWrap; a11.fill = greenFill; a11.border = thinBorder;
  const c11 = summarySheet.getCell("C11");
  c11.value = displayExactValue(totalNetInjectedForGreenBox);
  c11.font = valueFont; c11.alignment = centerMid; c11.fill = greenFill; c11.border = thinBorder;

  // D7..I12 detail strip (ABT/Voltage/CTPT/CT/PT/MF)
  const abtMeterLabelCell = summarySheet.getCell("D7");
  abtMeterLabelCell.value = "ABT Main Meter Sr. No.";
  abtMeterLabelCell.font = { bold: true, size: 10, name: "Times New Roman" };
  abtMeterLabelCell.alignment = { horizontal: "right", vertical: "middle", wrapText: true };
  abtMeterLabelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
  abtMeterLabelCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  const meterCells = [{ col: "E7", value: lossesCalculationData.mainClient.meterNumber || "", bgColor: "D9D9D9" }];
  const maxSubClients = lossesCalculationData.subClient.length;
  for (let i = 0; i < maxSubClients; i++) {
    const colChar = String.fromCharCode(69 + i + 1); // F onwards
    meterCells.push({
      col: `${colChar}7`,
      value: i < lossesCalculationData.subClient.length ? (lossesCalculationData.subClient[i].meterNumber || "") : "",
      bgColor: "D9D9D9",
    });
  }
  meterCells.forEach((m) => {
    const cell = summarySheet.getCell(m.col);
    cell.value = m.value;
    cell.font = { bold: true, size: 10, name: "Times New Roman" };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: m.bgColor } };
    cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
  });

  // Fill intermediate columns with gray background up to K column (row 7)
  // Minimum K column (75) tak gray background, but agar 6+ subclients hain to extra gray column nahi dikhana
  if (maxSubClients <= 5) {
    const lastSubClientCol = 70 + maxSubClients; // Column after last subclient
    const kColumn = 75; // K column
    // Fill all columns from last subclient to K with gray
    for (let colCode = lastSubClientCol; colCode <= kColumn; colCode++) {
      const col = String.fromCharCode(colCode);
      const cellRef = summarySheet.getCell(`${col}7`);
      cellRef.value = "";
      cellRef.font = { bold: true, size: 10, name: "Times New Roman" };
      cellRef.alignment = { horizontal: "center", vertical: "middle" };
      cellRef.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
      cellRef.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    }
  }

  const voltageLabelCell = summarySheet.getCell("D8");
  voltageLabelCell.value = "Voltage Level";
  voltageLabelCell.font = { bold: true, size: 10, name: "Times New Roman" };
  voltageLabelCell.alignment = { horizontal: "right", vertical: "middle", wrapText: true };
  voltageLabelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
  voltageLabelCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  const voltageCells = [{ col: "E8", value: lossesCalculationData.mainClient.mainClientDetail.voltageLevel || "", bgColor: "D9D9D9" }];
  for (let i = 0; i < maxSubClients; i++) {
    const colChar = String.fromCharCode(69 + i + 1);
    voltageCells.push({
      col: `${colChar}8`,
      value: i < lossesCalculationData.subClient.length ? (lossesCalculationData.subClient[i].voltageLevel || "") : "",
      bgColor: "D9D9D9",
    });
  }
  voltageCells.forEach((v) => {
    const cell = summarySheet.getCell(v.col);
    cell.value = v.value;
    cell.font = { size: 10, name: "Times New Roman" };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: v.bgColor } };
    cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
  });

  // Fill intermediate columns with gray background up to K column (row 8)
  // Minimum K column (75) tak gray background, but agar 6+ subclients hain to extra gray column nahi dikhana
  if (maxSubClients <= 5) {
    const lastSubClientCol = 70 + maxSubClients; // Column after last subclient
    const kColumn = 75; // K column
    // Fill all columns from last subclient to K with gray
    for (let colCode = lastSubClientCol; colCode <= kColumn; colCode++) {
      const col = String.fromCharCode(colCode);
      const cellRef = summarySheet.getCell(`${col}8`);
      cellRef.value = "";
      cellRef.font = { size: 10, name: "Times New Roman" };
      cellRef.alignment = { horizontal: "center", vertical: "middle" };
      cellRef.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
      cellRef.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    }
  }

  const ctptLabelCell = summarySheet.getCell("D9");
  ctptLabelCell.value = "CTPT Sr.No.";
  ctptLabelCell.font = { bold: true, size: 10, name: "Times New Roman" };
  ctptLabelCell.alignment = { horizontal: "right", vertical: "middle" };
  ctptLabelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
  ctptLabelCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  const ctptCells = [{ col: "E9", value: lossesCalculationData.mainClient.mainClientDetail.ctptSrNo || "", bgColor: "D9D9D9" }];
  for (let i = 0; i < maxSubClients; i++) {
    const colChar = String.fromCharCode(69 + i + 1);
    ctptCells.push({
      col: `${colChar}9`,
      value: i < lossesCalculationData.subClient.length ? (lossesCalculationData.subClient[i].ctptSrNo || "") : "",
      bgColor: "D9D9D9",
    });
  }
  ctptCells.forEach((v) => {
    const cell = summarySheet.getCell(v.col);
    cell.value = v.value;
    cell.font = { size: 10, name: "Times New Roman" };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: v.bgColor } };
    cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
  });

  // Fill intermediate columns with gray background up to K column (row 9)
  // Minimum K column (75) tak gray background, but agar 6+ subclients hain to extra gray column nahi dikhana
  if (maxSubClients <= 5) {
    const lastSubClientCol = 70 + maxSubClients; // Column after last subclient
    const kColumn = 75; // K column
    // Fill all columns from last subclient to K with gray
    for (let colCode = lastSubClientCol; colCode <= kColumn; colCode++) {
      const col = String.fromCharCode(colCode);
      const cellRef = summarySheet.getCell(`${col}9`);
      cellRef.value = "";
      cellRef.font = { size: 10, name: "Times New Roman" };
      cellRef.alignment = { horizontal: "center", vertical: "middle" };
      cellRef.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
      cellRef.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    }
  }

  const ctRatioLabelCell = summarySheet.getCell("D10");
  ctRatioLabelCell.value = "CT Ratio (A/A)";
  ctRatioLabelCell.font = { bold: true, size: 10, name: "Times New Roman" };
  ctRatioLabelCell.alignment = { horizontal: "right", vertical: "middle" };
  ctRatioLabelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
  ctRatioLabelCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  const ctRatioCells = [{ col: "E10", value: lossesCalculationData.mainClient.mainClientDetail.ctRatio || "", bgColor: "D9D9D9" }];
  for (let i = 0; i < maxSubClients; i++) {
    const colChar = String.fromCharCode(69 + i + 1);
    ctRatioCells.push({
      col: `${colChar}10`,
      value: i < lossesCalculationData.subClient.length ? (lossesCalculationData.subClient[i].ctRatio || "") : "",
      bgColor: "D9D9D9",
    });
  }
  ctRatioCells.forEach((v) => {
    const cell = summarySheet.getCell(v.col);
    cell.value = v.value;
    cell.font = { size: 10, name: "Times New Roman" };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: v.bgColor } };
    cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
  });

  // Fill intermediate columns with gray background up to K column (row 10)
  // Minimum K column (75) tak gray background, but agar 6+ subclients hain to extra gray column nahi dikhana
  if (maxSubClients <= 5) {
    const lastSubClientCol = 70 + maxSubClients; // Column after last subclient
    const kColumn = 75; // K column
    // Fill all columns from last subclient to K with gray
    for (let colCode = lastSubClientCol; colCode <= kColumn; colCode++) {
      const col = String.fromCharCode(colCode);
      const cellRef = summarySheet.getCell(`${col}10`);
      cellRef.value = "";
      cellRef.font = { size: 10, name: "Times New Roman" };
      cellRef.alignment = { horizontal: "center", vertical: "middle" };
      cellRef.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
      cellRef.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    }
  }

  const ptRatioLabelCell = summarySheet.getCell("D11");
  ptRatioLabelCell.value = "PT Ratio (V/V)";
  ptRatioLabelCell.font = { bold: true, size: 10, name: "Times New Roman" };
  ptRatioLabelCell.alignment = { horizontal: "right", vertical: "middle" };
  ptRatioLabelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
  ptRatioLabelCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  const ptRatioCells = [{ col: "E11", value: lossesCalculationData.mainClient.mainClientDetail.ptRatio || "", bgColor: "D9D9D9" }];
  for (let i = 0; i < maxSubClients; i++) {
    const colChar = String.fromCharCode(69 + i + 1);
    ptRatioCells.push({
      col: `${colChar}11`,
      value: i < lossesCalculationData.subClient.length ? (lossesCalculationData.subClient[i].ptRatio || "") : "",
      bgColor: "D9D9D9",
    });
  }
  ptRatioCells.forEach((v) => {
    const cell = summarySheet.getCell(v.col);
    cell.value = v.value;
    cell.font = { size: 10, name: "Times New Roman" };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: v.bgColor } };
    cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
  });

  // Fill intermediate columns with gray background up to K column (row 11)
  // Minimum K column (75) tak gray background, but agar 6+ subclients hain to extra gray column nahi dikhana
  if (maxSubClients <= 5) {
    const lastSubClientCol = 70 + maxSubClients; // Column after last subclient
    const kColumn = 75; // K column
    // Fill all columns from last subclient to K with gray
    for (let colCode = lastSubClientCol; colCode <= kColumn; colCode++) {
      const col = String.fromCharCode(colCode);
      const cellRef = summarySheet.getCell(`${col}11`);
      cellRef.value = "";
      cellRef.font = { size: 10, name: "Times New Roman" };
      cellRef.alignment = { horizontal: "center", vertical: "middle" };
      cellRef.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
      cellRef.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    }
  }

  const mfLabelCell = summarySheet.getCell("D12");
  mfLabelCell.value = "MF";
  mfLabelCell.font = { bold: true, size: 10, name: "Times New Roman" };
  mfLabelCell.alignment = { horizontal: "right", vertical: "middle" };
  mfLabelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
  mfLabelCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  const mfCells = [{
    col: "E12",
    value: lossesCalculationData.mainClient.mainClientDetail.mf || "5000",
    bgColor: "D9D9D9",
  }];
  for (let i = 0; i < maxSubClients; i++) {
    const colChar = String.fromCharCode(69 + i + 1);
    mfCells.push({
      col: `${colChar}12`,
      value: i < lossesCalculationData.subClient.length ? (lossesCalculationData.subClient[i].mf || "1000") : "",
      bgColor: "D9D9D9",
    });
  }
  mfCells.forEach((v) => {
    const cell = summarySheet.getCell(v.col);
    cell.value = v.value;
    cell.font = { size: 10, name: "Times New Roman" };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: v.bgColor } };
    cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
  });

  // Fill intermediate columns with gray background up to K column (row 12)
  // Minimum K column (75) tak gray background, but agar 6+ subclients hain to extra gray column nahi dikhana
  if (maxSubClients <= 5) {
    const lastSubClientCol = 70 + maxSubClients; // Column after last subclient
    const kColumn = 75; // K column
    // Fill all columns from last subclient to K with gray
    for (let colCode = lastSubClientCol; colCode <= kColumn; colCode++) {
      const col = String.fromCharCode(colCode);
      const cellRef = summarySheet.getCell(`${col}12`);
      cellRef.value = "";
      cellRef.font = { size: 10, name: "Times New Roman" };
      cellRef.alignment = { horizontal: "center", vertical: "middle" };
      cellRef.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
      cellRef.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    }
  }

  // Row 13 spacer before table - merged cell from A to last column
  summarySheet.getRow(13).height = 15;
  summarySheet.mergeCells(`A13:${lastMergedColumnChar}13`);
  const row13Cell = summarySheet.getCell("A13");
  row13Cell.value = "";
  row13Cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF" } };

  // === Rows 14–16: Overall Distributions to Distribution Licensee (DISCOM) ===
  // Left merged label block A14:B16
  summarySheet.mergeCells("A14:B16");
  const discomBlockLabel = summarySheet.getCell("A14");
  discomBlockLabel.value = "Overall Distributions to Distribution\nLicensee ( DISCOM )";
  discomBlockLabel.font = { bold: true, size: 11, name: "Times New Roman" };
  discomBlockLabel.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  discomBlockLabel.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF00" } };
  discomBlockLabel.border = {
    top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" },
  };

  // Row labels in column C (yellow)
  const r14 = summarySheet.getCell("C14");
  r14.value = "DISCOM Name";
  r14.font = { bold: true, size: 10, name: "Times New Roman" };
  r14.alignment = { horizontal: "center", vertical: "middle" };
  r14.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF00" } };
  r14.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  const r15 = summarySheet.getCell("C15");
  r15.value = "TOTAL Units";
  r15.font = { bold: true, size: 10, name: "Times New Roman" };
  r15.alignment = { horizontal: "center", vertical: "middle" };
  r15.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF00" } };
  r15.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  const r16 = summarySheet.getCell("C16");
  r16.value = "Percentage (%)";
  r16.font = { bold: true, size: 10, name: "Times New Roman" };
  r16.alignment = { horizontal: "center", vertical: "middle" };
  r16.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF00" } };
  r16.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  // DISCOM columns D..J
  const discoms = ["DGVCL", "MGVCL", "PGVCL", "UGVCL", "TAECO", "TSECO", "TEL", "TOTAL"];

  // Pull values from DB (null -> 0)
  const rawValues = [
    Number(lossesCalculationData.DGVCL || 0),
    Number(lossesCalculationData.MGVCL || 0),
    Number(lossesCalculationData.PGVCL || 0),
    Number(lossesCalculationData.UGVCL || 0),
    Number(lossesCalculationData.TAECO || 0),
    Number(lossesCalculationData.TSECO || 0),
    Number(lossesCalculationData.TEL || 0),
  ];

  // Compute total (8th column J)
  const totalUnits = rawValues.reduce((a, b) => a + b, 0);
  const values = [...rawValues, totalUnits];

  // helper to get column letter from index 0..7 -> D..J
  const colFromIndex = (i) => String.fromCharCode("D".charCodeAt(0) + i);

  // Colors
  const yellow = "FFFF00";
  const lightPink = "FFC7CE";

  // Build rows 14 (names), 15 (units), 16 (percentages)
  for (let i = 0; i < discoms.length; i++) {
    const col = colFromIndex(i);

    // Check if cell has data (non-zero) - used for all rows
    const hasData = values[i] > 0;

    // Row 14: DISCOM Name
    // If cell has data (non-zero), use pink background, otherwise yellow
    const nameCellBgColor = hasData ? lightPink : yellow;
    const nameCell = summarySheet.getCell(`${col}14`);
    nameCell.value = discoms[i];
    nameCell.font = { bold: true, size: 10, name: "Times New Roman" };
    nameCell.alignment = { horizontal: "center", vertical: "middle" };
    nameCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: nameCellBgColor } };
    nameCell.border = {
      top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" },
      right: { style: i === discoms.length - 1 ? "thin" : "thin" },
    };

    // Row 15: TOTAL Units (3-decimals as text)
    // If cell has data (non-zero), use pink background, otherwise yellow
    const valCellBgColor = hasData ? lightPink : yellow;
    const valCell = summarySheet.getCell(`${col}15`);
    valCell.value = displayExactValue(values[i]); // uses your existing helper
    valCell.font = { bold: true, size: 10, name: "Times New Roman" };
    valCell.alignment = { horizontal: "center", vertical: "middle" };
    valCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: valCellBgColor } };
    valCell.border = {
      top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" },
      right: { style: i === discoms.length - 1 ? "thin" : "thin" },
    };

    // Row 16: Percentage (2-decimals + %)
    // If cell has data (non-zero), use pink background, otherwise yellow
    const pct = totalUnits > 0 ? (values[i] / totalUnits) * 100 : 0;
    const pctCellBgColor = hasData ? lightPink : yellow;
    const pctCell = summarySheet.getCell(`${col}16`);
    pctCell.value = `${pct.toFixed(2)}%`;
    pctCell.font = { bold: true, size: 10, name: "Times New Roman" };
    pctCell.alignment = { horizontal: "center", vertical: "middle" };
    pctCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: pctCellBgColor } };
    pctCell.border = {
      top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" },
      right: { style: i === discoms.length - 1 ? "thin" : "thin" },
    };
  }

  // Add blank yellow cells after TOTAL column (K) up to last subclient column for rows 14-16
  // TOTAL is at column K (75, index 7 in discoms array), so we need to fill from L (76) onwards
  const totalColCode = "K".charCodeAt(0); // 75 (TOTAL column)
  const startColCode = totalColCode + 1; // L = 76 (after TOTAL)
  if (lastMergedColumn >= startColCode) {
    for (let colCode = startColCode; colCode <= lastMergedColumn; colCode++) {
      const col = String.fromCharCode(colCode);

      // Row 14: Blank yellow cell
      const blankCell14 = summarySheet.getCell(`${col}14`);
      blankCell14.value = "";
      blankCell14.font = { bold: true, size: 10, name: "Times New Roman" };
      blankCell14.alignment = { horizontal: "center", vertical: "middle" };
      blankCell14.fill = { type: "pattern", pattern: "solid", fgColor: { argb: yellow } };
      blankCell14.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

      // Row 15: Blank yellow cell
      const blankCell15 = summarySheet.getCell(`${col}15`);
      blankCell15.value = "";
      blankCell15.font = { bold: true, size: 10, name: "Times New Roman" };
      blankCell15.alignment = { horizontal: "center", vertical: "middle" };
      blankCell15.fill = { type: "pattern", pattern: "solid", fgColor: { argb: yellow } };
      blankCell15.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

      // Row 16: Blank yellow cell
      const blankCell16 = summarySheet.getCell(`${col}16`);
      blankCell16.value = "";
      blankCell16.font = { bold: true, size: 10, name: "Times New Roman" };
      blankCell16.alignment = { horizontal: "center", vertical: "middle" };
      blankCell16.fill = { type: "pattern", pattern: "solid", fgColor: { argb: yellow } };
      blankCell16.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    }
  }

  // Make row heights match your style
  summarySheet.getRow(14).height = 28;
  summarySheet.getRow(15).height = 28;
  summarySheet.getRow(16).height = 28;

  // Row 17 spacer before table - merged cell from A to last column
  summarySheet.getRow(17).height = 15;
  summarySheet.mergeCells(`A17:${lastMergedColumnChar}17`);
  const row17Cell = summarySheet.getCell("A17");
  row17Cell.value = "";
  row17Cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF" } };

  // =========================
  // SHIFTED TABLE START HERE
  // =========================
  const tableStartRow = 18;           // (was 14) → moved entire table down by 4 rows
  const headerRow = tableStartRow;    // 18
  const dataStartRow = headerRow + 1; // 19

  // Header row
  const tableHeaderRow = summarySheet.getRow(headerRow);
  tableHeaderRow.height = 54;

  const tableHeaders = [
    { cell: `A${headerRow}`, value: "Sr. No.", bgColor: "D9D9D9", borderRight: "thin", borderLeft: "thin" },
    { cell: `B${headerRow}`, value: "HT Consumer Name", bgColor: "D9D9D9" },
    { cell: `C${headerRow}`, value: "HT / LTMD Consumer No.", bgColor: "D9D9D9" },
    { cell: `D${headerRow}`, value: "Wheeling Division Office/Location", bgColor: "D9D9D9" },
    { cell: `E${headerRow}`, value: "Wheeling DISCOM", bgColor: "D9D9D9" },
    { cell: `F${headerRow}`, value: "Project Capacity (kW) (AC)", bgColor: "D9D9D9", borderRight: "thin" },
    { cell: `G${headerRow}`, value: "Share in Gross Injected Units to S/S", bgColor: "D9D9D9" },
    { cell: `H${headerRow}`, value: "Share in Gross Drawl Units from S/S", bgColor: "D9D9D9" },
    { cell: `I${headerRow}`, value: "Net Injected Units to S/S", bgColor: "D9D9D9" },
    { cell: `J${headerRow}`, value: "% Weightage According to Gross Injecting", bgColor: "D9D9D9" },
  ];

  tableHeaders.forEach((h) => {
    const cell = summarySheet.getCell(h.cell);
    cell.value = h.value;
    cell.font = { bold: true, size: 10, name: "Times New Roman" };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: h.bgColor } };
    cell.border = {
      top: { style: "thin" },
      left: { style: h.borderLeft || "thin" },
      bottom: { style: "thin" },
      right: { style: h.borderRight || "thin" },
    };
  });

  // Remark column header - extend dynamically based on subclients
  // Minimum K (75), but if subclients > 5, extend to last subclient column
  const remarkStartCol = 75; // K
  const remarkEndCol = lastMergedColumn; // Dynamic based on subclients
  summarySheet.mergeCells(`K${headerRow}:${lastMergedColumnChar}${headerRow}`);
  const k18Cell = summarySheet.getCell(`K${headerRow}`);
  k18Cell.value = "Remark";
  k18Cell.font = { bold: true, size: 10, name: "Times New Roman" };
  k18Cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  k18Cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
  k18Cell.border = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };

  // Outer thin border for header row
  const firstHeaderCol = "A";
  const leftCell = summarySheet.getCell(`${firstHeaderCol}${headerRow}`);
  leftCell.border = { ...leftCell.border, left: { style: "thin" } };
  const rightCell = summarySheet.getCell(`J${headerRow}`);
  rightCell.border = { ...rightCell.border, right: { style: "thin" } };
  for (let col = 1; col <= 10; col++) {
    const colChar = String.fromCharCode(64 + col);
    const cell = summarySheet.getCell(`${colChar}${headerRow}`);
    cell.border = { ...cell.border, top: { style: "thin" }, bottom: { style: "thin" } };
  }
  // Add thin top and bottom borders to Remark column (merged K to lastMergedColumn)
  k18Cell.border = { ...k18Cell.border, top: { style: "thin" }, bottom: { style: "thin" } };

  // Data rows
  let rowNum = dataStartRow; // starts at 19
  let globalIndex = 1;

  lossesCalculationData.subClient.forEach((subClient) => {
    const subClientData = subClient.subClientsData || {};

    if (subClientData.partclient && subClientData.partclient.length > 0) {
      subClientData.partclient.forEach((partClient, partIndex) => {
        const currentRowNum = rowNum++;
        const row = summarySheet.getRow(currentRowNum);
        row.height = 60;

        const grossInjection = partClient.grossInjectionMWHAfterLosses || 0;
        const drawl = partClient.drawlMWHAfterLosses || 0;
        const netInjection = grossInjection + drawl;
        const weightage = totalGrossInjectedForGreenBox > 0 ? (grossInjection / totalGrossInjectedForGreenBox) * 100 : 0;

        // A: Sr No
        summarySheet.getCell(`A${currentRowNum}`).value = globalIndex++;
        summarySheet.getCell(`A${currentRowNum}`).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        summarySheet.getCell(`A${currentRowNum}`).border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

        // B: Name + sharing
        summarySheet.getCell(`B${currentRowNum}`).value = {
          richText: [
            { text: `${subClient.name.toUpperCase()} - Unit-${partIndex + 1}`, font: { size: 10, name: "Times New Roman" } },
            { text: ` (${partClient.sharingPercentage}% of Sharing OA)`, font: { size: 10, name: "Times New Roman" } },
          ],
        };
        summarySheet.getCell(`B${currentRowNum}`).alignment = { horizontal: "left", vertical: "middle", wrapText: true };
        summarySheet.getCell(`B${currentRowNum}`).border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
        summarySheet.getCell(`B${currentRowNum}`).font = { size: 10, name: "Times New Roman" };

        // C..E
        summarySheet.getCell(`C${currentRowNum}`).value = partClient.consumerNo || "";
        summarySheet.getCell(`C${currentRowNum}`).alignment = { horizontal: "center", vertical: "middle" };

        summarySheet.getCell(`D${currentRowNum}`).value = subClient.divisionName || "";
        summarySheet.getCell(`D${currentRowNum}`).alignment = { horizontal: "center", vertical: "middle", wrapText: true };

        summarySheet.getCell(`E${currentRowNum}`).value = subClient.discom || "";
        summarySheet.getCell(`E${currentRowNum}`).alignment = { horizontal: "center", vertical: "middle" };

        // F: Capacity (merge for partclients)
        if (partIndex === 0) {
          summarySheet.getCell(`F${currentRowNum}`).value = subClient.acCapacityKw || "";
          summarySheet.getCell(`F${currentRowNum}`).alignment = { horizontal: "center", vertical: "middle" };
          summarySheet.getCell(`F${currentRowNum}`).border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

          if (subClientData.partclient.length > 1) {
            summarySheet.mergeCells(`F${currentRowNum}:F${currentRowNum + subClientData.partclient.length - 1}`);
            for (let i = 0; i < subClientData.partclient.length; i++) {
              summarySheet.getCell(`F${currentRowNum + i}`).border = {
                top: i === 0 ? { style: "thin" } : { style: "none" },
                left: { style: "thin" },
                bottom: i === subClientData.partclient.length - 1 ? { style: "thin" } : { style: "none" },
                right: { style: "thin" },
              };
            }
          }
        }

        // Base formatting for other cells (except B,F done above)
        for (let col = 3; col <= 10; col++) {
          if (col !== 6) {
            const cell = row.getCell(col);
            cell.font = { size: 10, name: "Times New Roman" };
            cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
            cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
          }
        }

        // Bold numeric cols
        ["G", "H", "I", "J"].forEach((col) => {
          const cell = summarySheet.getCell(`${col}${currentRowNum}`);
          cell.font = { ...cell.font, bold: true };
        });

        summarySheet.getCell(`G${currentRowNum}`).value = displayExactValue(grossInjection);
        summarySheet.getCell(`H${currentRowNum}`).value = displayExactValue(drawl);
        summarySheet.getCell(`I${currentRowNum}`).value = displayExactValue(netInjection);
        summarySheet.getCell(`J${currentRowNum}`).value = `${Number(weightage).toFixed(2)} %`;

        // Remark column (K to lastMergedColumn) - blank white cells for data rows
        for (let colCode = remarkStartCol; colCode <= remarkEndCol; colCode++) {
          const col = String.fromCharCode(colCode);
          const remarkCell = summarySheet.getCell(`${col}${currentRowNum}`);
          remarkCell.value = "";
          remarkCell.font = { size: 10, name: "Times New Roman" };
          remarkCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
          remarkCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF" } }; // White background
          remarkCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
        }

        // Increase border for rows 25-27, columns F to lastMergedColumn
        if (currentRowNum >= 25 && currentRowNum <= 27) {
          const colsFtoJ = ["F", "G", "H", "I", "J"];
          colsFtoJ.forEach((col) => {
            const cell = summarySheet.getCell(`${col}${currentRowNum}`);
            const currentBorder = cell.border || {};
            cell.border = {
              top: { style: "thin" },
              left: { style: currentBorder.left?.style === "thin" ? "thin" : "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          });
          // Also update remark columns
          for (let colCode = remarkStartCol; colCode <= remarkEndCol; colCode++) {
            const col = String.fromCharCode(colCode);
            const cell = summarySheet.getCell(`${col}${currentRowNum}`);
            cell.border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }
        }
      });
    } else {
      // single subclient row
      const currentRowNum = rowNum++;
      const row = summarySheet.getRow(currentRowNum);
      row.height = 60;

      const grossInjection = subClientData.grossInjectionMWHAfterLosses || 0;
      const drawl = subClientData.drawlMWHAfterLosses || 0;
      const netInjection = subClientData.netInjectionMWHAfterLosses || 0;
      const weightage = totalGrossInjectedForGreenBox > 0 ? (grossInjection / totalGrossInjectedForGreenBox) * 100 : 0;

      summarySheet.getCell(`A${currentRowNum}`).value = globalIndex++;
      summarySheet.getCell(`A${currentRowNum}`).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      summarySheet.getCell(`A${currentRowNum}`).border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

      summarySheet.getCell(`B${currentRowNum}`).value = `${subClient.name.toUpperCase()}`;
      summarySheet.getCell(`B${currentRowNum}`).alignment = { horizontal: "left", vertical: "middle", wrapText: true };
      summarySheet.getCell(`B${currentRowNum}`).border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      summarySheet.getCell(`B${currentRowNum}`).font = { size: 10, name: "Times New Roman" };

      summarySheet.getCell(`C${currentRowNum}`).value = subClient.consumerNo || "";
      summarySheet.getCell(`C${currentRowNum}`).alignment = { horizontal: "center", vertical: "middle" };

      summarySheet.getCell(`D${currentRowNum}`).value = subClient.divisionName || "";
      summarySheet.getCell(`D${currentRowNum}`).alignment = { horizontal: "center", vertical: "middle", wrapText: true };

      summarySheet.getCell(`E${currentRowNum}`).value = subClient.discom;
      summarySheet.getCell(`E${currentRowNum}`).alignment = { horizontal: "center", vertical: "middle" };

      summarySheet.getCell(`F${currentRowNum}`).value = subClient.acCapacityKw || "";
      summarySheet.getCell(`F${currentRowNum}`).alignment = { horizontal: "center", vertical: "middle" };
      summarySheet.getCell(`F${currentRowNum}`).border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

      for (let col = 3; col <= 10; col++) {
        if (col !== 6) {
          const cell = row.getCell(col);
          cell.font = { size: 10, name: "Times New Roman" };
          cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
          cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
        }
      }

      ["G", "H", "I", "J"].forEach((col) => {
        const cell = summarySheet.getCell(`${col}${currentRowNum}`);
        cell.font = { ...cell.font, bold: true };
      });

      summarySheet.getCell(`G${currentRowNum}`).value = displayExactValue(grossInjection);
      summarySheet.getCell(`H${currentRowNum}`).value = displayExactValue(drawl);
      summarySheet.getCell(`I${currentRowNum}`).value = displayExactValue(netInjection);
      summarySheet.getCell(`J${currentRowNum}`).value = `${Number(weightage).toFixed(2)} %`;

      // Remark column (K to lastMergedColumn) - blank white cells for data rows
      for (let colCode = remarkStartCol; colCode <= remarkEndCol; colCode++) {
        const col = String.fromCharCode(colCode);
        const remarkCell = summarySheet.getCell(`${col}${currentRowNum}`);
        remarkCell.value = "";
        remarkCell.font = { size: 10, name: "Times New Roman" };
        remarkCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        remarkCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF" } }; // White background
        remarkCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      }

      // Increase border for rows 25-27, columns F to lastMergedColumn
      if (currentRowNum >= 25 && currentRowNum <= 27) {
        const colsFtoJ = ["F", "G", "H", "I", "J"];
        colsFtoJ.forEach((col) => {
          const cell = summarySheet.getCell(`${col}${currentRowNum}`);
          const currentBorder = cell.border || {};
          cell.border = {
            top: { style: "thin" },
            left: { style: currentBorder.left?.style === "thin" ? "thin" : "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        });
        // Also update remark columns
        for (let colCode = remarkStartCol; colCode <= remarkEndCol; colCode++) {
          const col = String.fromCharCode(colCode);
          const cell = summarySheet.getCell(`${col}${currentRowNum}`);
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        }
      }
    }
  });

  // Blank row after data
  const blankRowNum = rowNum++;
  const blankRow = summarySheet.getRow(blankRowNum);
  blankRow.height = 15;
  for (let col = 1; col <= 10; col++) {
    const cell = blankRow.getCell(col);
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: col === 6 ? "thin" : "thin" }, // keep F thin border
    };
  }
  // Remark column cells for blank row
  for (let colCode = remarkStartCol; colCode <= remarkEndCol; colCode++) {
    const col = String.fromCharCode(colCode);
    const blankCell = summarySheet.getCell(`${col}${blankRowNum}`);
    blankCell.value = "";
    blankCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
  }
  // Increase border for rows 25-27, columns F to lastMergedColumn (blank row)
  if (blankRowNum >= 25 && blankRowNum <= 27) {
    const colsFtoJ = ["F", "G", "H", "I", "J"];
    colsFtoJ.forEach((col) => {
      const cell = summarySheet.getCell(`${col}${blankRowNum}`);
      cell.border = {
        top: { style: "thin" },
        left: { style: col === "F" ? "thin" : "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });
    // Also update remark columns
    for (let colCode = remarkStartCol; colCode <= remarkEndCol; colCode++) {
      const col = String.fromCharCode(colCode);
      const cell = summarySheet.getCell(`${col}${blankRowNum}`);
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    }
  }

  // Totals row (immediately after blank)
  const totalRowNum = blankRowNum + 1;
  const totalRow = summarySheet.getRow(totalRowNum);
  totalRow.height = 28;

  summarySheet.mergeCells(`A${totalRowNum}:E${totalRowNum}`);
  const totalLabelCells = summarySheet.getCell(`A${totalRowNum}`);
  totalLabelCells.value = "Total";
  totalLabelCells.font = { bold: true, size: 12, name: "Times New Roman" };
  totalLabelCells.alignment = { horizontal: "center", vertical: "middle" };
  totalLabelCells.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
  totalLabelCells.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  // Compute totals
  let totalCapacity = 0;
  let totalGrossInjected = 0;
  let totalGrossDrawl = 0;

  lossesCalculationData.subClient.forEach((subClient) => {
    const sc = subClient.subClientsData || {};
    if (sc.partclient && sc.partclient.length > 0) {
      totalCapacity += subClient.acCapacityKw || 0;
      sc.partclient.forEach((pc) => {
        totalGrossInjected += pc.grossInjectionMWHAfterLosses || 0;
        totalGrossDrawl += pc.drawlMWHAfterLosses || 0;
      });
    } else {
      totalCapacity += subClient.acCapacityKw || 0;
      totalGrossInjected += sc.grossInjectionMWHAfterLosses || 0;
      totalGrossDrawl += sc.drawlMWHAfterLosses || 0;
    }
  });

  const totalNetInjected = totalGrossInjected + totalGrossDrawl;

  // F..J totals
  const fCell = summarySheet.getCell(`F${totalRowNum}`);
  fCell.value = totalCapacity;
  fCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  summarySheet.getCell(`G${totalRowNum}`).value = displayExactValue(totalGrossInjected);
  summarySheet.getCell(`H${totalRowNum}`).value = displayExactValue(totalGrossDrawl);
  summarySheet.getCell(`I${totalRowNum}`).value = displayExactValue(totalNetInjected);
  summarySheet.getCell(`J${totalRowNum}`).value = "100%";

  for (let col = 6; col <= 10; col++) {
    const cell = totalRow.getCell(col);
    cell.font = { bold: true, size: 12, name: "Times New Roman" };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: col === 6 ? "thin" : "thin" },
    };
  }
  // Remark column cells for totals row - gray background
  for (let colCode = remarkStartCol; colCode <= remarkEndCol; colCode++) {
    const col = String.fromCharCode(colCode);
    const totalCell = summarySheet.getCell(`${col}${totalRowNum}`);
    totalCell.value = "";
    totalCell.font = { bold: true, size: 12, name: "Times New Roman" };
    totalCell.alignment = { horizontal: "center", vertical: "middle" };
    totalCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } }; // Gray background
    totalCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
  }

  // Increase border for rows 25-27, columns F to lastMergedColumn (totals row)
  if (totalRowNum >= 25 && totalRowNum <= 27) {
    const colsFtoJ = ["F", "G", "H", "I", "J"];
    colsFtoJ.forEach((col) => {
      const cell = summarySheet.getCell(`${col}${totalRowNum}`);
      cell.border = {
        top: { style: "thin" },
        left: { style: col === "F" ? "thin" : "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });
    // Also update remark columns
    for (let colCode = remarkStartCol; colCode <= remarkEndCol; colCode++) {
      const col = String.fromCharCode(colCode);
      const cell = summarySheet.getCell(`${col}${totalRowNum}`);
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    }
  }

  // Spacer after totals - extend to last column
  const spacerRowNum = totalRowNum + 1;
  summarySheet.getRow(spacerRowNum).height = 15;
  summarySheet.mergeCells(`A${spacerRowNum}:${lastMergedColumnChar}${spacerRowNum}`);
  const spacerCell = summarySheet.getCell(`A${spacerRowNum}`);
  spacerCell.value = "";
  spacerCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF" } };

  // Notes (shifted down automatically)
  const noteRow1 = totalRowNum + 2;
  const noteRow2 = totalRowNum + 3;

  summarySheet.getCell(`A${noteRow1}`).value = "Note:";
  summarySheet.getCell(`A${noteRow1}`).font = { bold: true, italic: true, size: 10, name: "Times New Roman" };
  summarySheet.mergeCells(`B${noteRow1}:${lastMergedColumnChar}${noteRow1}`);
  const noteCell1 = summarySheet.getCell(`B${noteRow1}`);
  noteCell1.value = {
    richText: [
      { text: "1) All Units are in ", font: { bold: true, italic: true, size: 10, name: "Times New Roman" } },
      { text: "MWH", font: { bold: true, italic: true, size: 10, name: "Times New Roman", color: { argb: "FF0000" } } },
    ],
  };

  summarySheet.getCell(`A${noteRow2}`).value = "";
  summarySheet.mergeCells(`B${noteRow2}:${lastMergedColumnChar}${noteRow2}`);
  const noteCell2 = summarySheet.getCell(`B${noteRow2}`);
  noteCell2.value = {
    richText: [
      { text: "2) ", font: { bold: true, italic: true, size: 10, name: "Times New Roman" } },
      { text: `${lossesCalculationData.mainClient.meterNumber}`, font: { italic: true, size: 10, name: "Times New Roman", bold: true } },
      { text: ` is the Grossing Meter at ${lossesCalculationData.mainClient.mainClientDetail.subTitle} S/S End`, font: { bold: true, italic: true, size: 10, name: "Times New Roman" } },
    ],
  };

  // Apply common formatting to both rows
  [noteRow1, noteRow2].forEach(row => {
    for (let col = 1; col <= 11; col++) {
      const cell = summarySheet.getCell(`${String.fromCharCode(64 + col)}${row}`);
      cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
  });

  // Add bold borders around the entire data area (from row 2 to note rows)
  const firstDataRow = 2;
  const lastDataRow = noteRow2; // Your last note row

  // Apply bold borders to the outer perimeter
  for (let row = firstDataRow; row <= lastDataRow; row++) {
    // Left border (column A)
    const leftCell = summarySheet.getCell(`A${row}`);
    leftCell.border = {
      ...leftCell.border,
      left: { style: 'thin' }
    };

    // Right border (column J) - exclude rows 13, 17, and 24 (spacer rows)
    if (row !== 13 && row !== 17 && row !== 24) {
      const rightCell = summarySheet.getCell(`J${row}`);
      rightCell.border = {
        ...rightCell.border,
        right: { style: 'thin' }
      };
    }

    // Right border (column K) for note rows (since notes extend to column K)
    if (row === noteRow1 || row === noteRow2) {
      const rightCellK = summarySheet.getCell(`K${row}`);
      rightCellK.border = {
        ...rightCellK.border,
        right: { style: 'thin' }
      };
    }
  }

  // Apply top border to first row
  for (let col = 1; col <= 10; col++) {
    const colChar = String.fromCharCode(64 + col);
    const cell = summarySheet.getCell(`${colChar}${firstDataRow}`);
    cell.border = {
      ...cell.border,
      top: { style: 'thin' }
    };
  }

  // Apply bottom border to last row
  for (let col = 1; col <= 10; col++) {
    const colChar = String.fromCharCode(64 + col);
    const cell = summarySheet.getCell(`${colChar}${lastDataRow}`);
    cell.border = {
      ...cell.border,
      bottom: { style: 'thin' }
    };
  }

  // Add complete outer border around entire table (A2 to K{lastDataRow})
  // Top border: A2 to K2
  for (let col = 1; col <= 11; col++) {
    const colChar = String.fromCharCode(64 + col);
    const topCell = summarySheet.getCell(`${colChar}${firstDataRow}`);
    topCell.border = {
      ...topCell.border,
      top: { style: 'thin' }
    };
  }

  // Bottom border: A{lastDataRow} to K{lastDataRow}
  for (let col = 1; col <= 11; col++) {
    const colChar = String.fromCharCode(64 + col);
    const bottomCell = summarySheet.getCell(`${colChar}${lastDataRow}`);
    bottomCell.border = {
      ...bottomCell.border,
      bottom: { style: 'thin' }
    };
  }

  // Left border: A2 to A{lastDataRow}
  for (let row = firstDataRow; row <= lastDataRow; row++) {
    const leftCell = summarySheet.getCell(`A${row}`);
    leftCell.border = {
      ...leftCell.border,
      left: { style: 'thin' }
    };
  }

  // Right border: K2 to K{lastDataRow} (complete outer border)
  for (let row = firstDataRow; row <= lastDataRow; row++) {
    const rightCell = summarySheet.getCell(`K${row}`);
    rightCell.border = {
      ...rightCell.border,
      right: { style: 'thin' }
    };
  }

  // Special handling for merged cells to ensure borders are visible
  const mergedAreas = [
    'A2:C2', 'D2:J2',   // Title row
    'A3:C3', 'D3:J3',   // Month row
    'A4:C4', 'D4:J4',   // Generation period
    'A5:J5',            // CPP Clients
    'A6:B6', 'A7:B7', 'A8:B8', 'A9:B9', 'A10:C10', 'A11:B11', 'A12:B12', // Left merged cells
    'A14:J14',          // Table headers
    `A${totalRowNum}:E${totalRowNum}`, // Total label
    `B${noteRow1}:K${noteRow1}`, `B${noteRow2}:K${noteRow2}` // Note rows
  ];

  mergedAreas.forEach(merge => {
    const cell = summarySheet.getCell(merge.split(':')[0]);
    cell.border = {
      ...cell.border,
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  });


  // Create a second sheet for masterdata with DGVCL format
  const masterdataSheet = workbook.addWorksheet('Master Data for DISCOM');
  // For the Master Data sheet
  masterdataSheet.pageSetup = worksheetSetup;

  // Set tab color (using exceljs)
  masterdataSheet.properties.tabColor = {
    argb: '92D050' // This is green color in ARGB format (Alpha, Red, Green, Blue)
  };

  masterdataSheet.pageSetup.orientation = 'landscape'; // Set page orientation to landscape
  masterdataSheet.pageSetup.fitToPage = true; // Fit to page
  masterdataSheet.pageSetup.fitToHeight = 1; // Fit to height
  masterdataSheet.pageSetup.fitToWidth = 1; // Fit to width

  // ===== TITLE SECTION =====
  // Add blank row at the top (row 1)
  masterdataSheet.insertRow(1);

  // Set column widths for Master Data sheet
  masterdataSheet.columns = [
    { width: 15 }, // A - Date
    { width: 12 }, // B - Block Time
    { width: 12 }, // C - Block No
    { width: 22 }, // D - Meter Number
    // Add more columns as needed for subclients
    // These will be set dynamically in the subclient loop
  ];

  // Note line (row 2)
  masterdataSheet.mergeCells('A2:J2');
  const noteCell = masterdataSheet.getCell('A2');

  noteCell.value = {
    richText: [
      { text: 'Note:- All Units are in ', font: { italic: true, bold: true, size: 12, name: 'Times New Roman', color: { argb: 'FF000000' } } },
      { text: 'MWH', font: { italic: true, bold: true, size: 12, name: 'Times New Roman', color: { argb: 'FFFF0000' } } } // Red color
    ]
  };

  noteCell.alignment = { horizontal: 'left', vertical: 'middle' };

  // Set row height for row 3
  masterdataSheet.getRow(3).height = 22;

  // Merge A3:F3 for company title
  masterdataSheet.mergeCells('A3:F3');
  const companyCell = masterdataSheet.getCell('A3');

  const acCapacityMw = (lossesCalculationData.mainClient.mainClientDetail.acCapacityKw / 1000).toFixed(2);
  companyCell.value = `${lossesCalculationData.mainClient.mainClientDetail.name.toUpperCase()} - ${acCapacityMw} MW AC Generation Details`;
  companyCell.font = { bold: true, size: 14, name: 'Times New Roman' };
  companyCell.alignment = { horizontal: 'left', vertical: 'middle' };
  companyCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF00' } // Yellow background
  };
  companyCell.border = {
    top: { style: 'thin', color: { argb: '000000' } },
    left: { style: 'thin', color: { argb: '000000' } },
    bottom: { style: 'thin', color: { argb: '000000' } },
    right: { style: 'thin', color: { argb: '000000' } }
  };

  // Month cell in G3
  const monthCell = masterdataSheet.getCell('G3');
  monthCell.value = `${monthName}-${lossesCalculationData.year.toString().slice(-2)}`;
  monthCell.font = {
    bold: true,
    size: 14,
    name: 'Times New Roman',
    color: { argb: 'FF0000' } // Red color
  };
  monthCell.alignment = { horizontal: 'center', vertical: 'middle' };
  monthCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF00' } // Yellow background
  };
  monthCell.border = {
    top: { style: 'thin', color: { argb: '000000' } },
    left: { style: 'thin', color: { argb: '000000' } },
    bottom: { style: 'thin', color: { argb: '000000' } },
    right: { style: 'thin', color: { argb: '000000' } }
  };

  // Date range from H3:I3 (will be extended later to include CHECK-SUM columns)
  masterdataSheet.mergeCells('H3:I3');
  const dateRangeCell = masterdataSheet.getCell('H3');
  dateRangeCell.value = `01-${monthStr}-${lossesCalculationData.year} to ${lastDay}-${monthStr}-${lossesCalculationData.year}`;
  dateRangeCell.font = {
    bold: true,
    size: 14,
    name: 'Times New Roman',
    color: { argb: 'FF0000' } // Red color
  };
  dateRangeCell.alignment = { horizontal: 'center', vertical: 'middle' };
  dateRangeCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF00' } // Yellow background
  };
  dateRangeCell.border = {
    top: { style: 'thin', color: { argb: '000000' } },
    left: { style: 'thin', color: { argb: '000000' } },
    bottom: { style: 'thin', color: { argb: '000000' } },
    right: { style: 'thin', color: { argb: '000000' } }
  };

  masterdataSheet.getColumn('H').width = 14;
  masterdataSheet.getColumn('I').width = 14;

  // Blank row (row 4)
  masterdataSheet.getRow(4).height = 15; // Slightly more spacing than your original 5
  // ===== MAIN DATA TABLE =====
  // Define the color scheme (same as ABT METER section)
  const clientColors = [
    'FFC000', // Orange
    'B4C6E7', // Light blue
    'E6B9D8', // Light purple
    'F8CBAD', // Peach
    'C6E0B4', // Light green
    'D9D9D9', // Light gray
    'E2EFDA', // Very light green
    'BDD7EE', // Light blue
    'FFF2CC', // Light yellow
    'DDEBF7'  // Very light blue
  ];
  // Headers
  // Merge A5:C5 for SLDC RAW DATA
  masterdataSheet.mergeCells('A5:C5');
  const sldcCell = masterdataSheet.getCell('A5');
  sldcCell.value = 'SLDC RAW DATA';
  sldcCell.font = { bold: true, size: 11, name: 'Times New Roman' };
  sldcCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
  sldcCell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };
  masterdataSheet.getRow(5).height = 62;

  const labels = ['WHEELING DISCOM', 'DIVISION NAME', 'CONSUMER NO.'];
  labels.forEach((label, index) => {
    const rowNumber = 6 + index;
    masterdataSheet.mergeCells(`A${rowNumber}:C${rowNumber}`);
    const cell = masterdataSheet.getCell(`A${rowNumber}`);
    cell.value = label;
    cell.font = { bold: true, size: 10, name: 'Times New Roman' };
    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
    masterdataSheet.getRow(rowNumber).height = 20;
  });

  // Insert blank row between CONSUMER NO. and GROSS INJECTION
  masterdataSheet.mergeCells('A9:C9');
  const gapLabelCell = masterdataSheet.getCell('A9');
  gapLabelCell.value = '';
  gapLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFFF' }
  };
  gapLabelCell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  // Merge D5:D8 for COMBINED meter header
  masterdataSheet.mergeCells('D5:D8');
  const combinedCell = masterdataSheet.getCell('D5');
  combinedCell.value = 'COMBINED (SS Side ABT METER Data)';
  combinedCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  combinedCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  combinedCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '92D050' }
  };
  combinedCell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };
  masterdataSheet.getColumn('D').width = 22;

  // Create headers for each subclient/partclient from column E onward
  let currentCol = 'E'; // Start from column E
  const subClientColumns = {};

  subClients.forEach((subClient, subIndex) => {
    const color = clientColors[subIndex % clientColors.length];
    const partClients = subClient.subClientsData?.partclient || [];

    if (partClients.length > 0) {
      // If subclient has part clients, create columns for each part client
      partClients.forEach((partClient, partIndex) => {
        const colLetter = currentCol;
        subClientColumns[`${subClient.name}-${partClient.divisionName}`] = colLetter;

        // Row 5 - Part Client name with percentage
        const cell1 = masterdataSheet.getCell(`${colLetter}5`);
        cell1.value = `${subClient.name.toUpperCase()} - Unit-${partIndex + 1} (${partClient.sharingPercentage}%)`;
        cell1.font = { bold: true, size: 10, name: 'Times New Roman' }; // Smaller font for long names
        cell1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell1.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: color } // Same color for all part clients of this subclient
        };
        cell1.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };

        // Row 6 - DISCOM
        const cell2 = masterdataSheet.getCell(`${colLetter}6`);
        masterdataSheet.getRow(6).height = 25;
        cell2.value = subClient.discom || '';
        cell2.font = { bold: true, size: 10, name: 'Times New Roman' };
        cell2.alignment = { horizontal: 'center', vertical: 'middle' };
        cell2.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: color }
        };
        cell2.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };

        // Row 7 - DIVISION NAME
        const cell3 = masterdataSheet.getCell(`${colLetter}7`);
        masterdataSheet.getRow(7).height = 45;
        cell3.value = partClient.divisionName || subClient.divisionName;
        cell3.font = { bold: true, size: 10, name: 'Times New Roman' };
        cell3.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell3.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: color }
        };
        cell3.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };

        // Row 8 - CONSUMER NO.
        const cell4 = masterdataSheet.getCell(`${colLetter}8`);
        masterdataSheet.getRow(8).height = 25;
        cell4.value = partClient.consumerNo || subClient.consumerNo;
        cell4.font = { bold: true, size: 10, name: 'Times New Roman' };
        cell4.alignment = { horizontal: 'center', vertical: 'middle' };
        cell4.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: color }
        };
        cell4.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };

        // Row 9 - Blank Row (color-matched to subclient)
        const cellGap = masterdataSheet.getCell(`${colLetter}9`);
        cellGap.value = '';
        cellGap.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: color }
        };
        cellGap.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };

        masterdataSheet.getColumn(colLetter).width = 22;
        currentCol = String.fromCharCode(currentCol.charCodeAt(0) + 1); // Move to next column
      });
    } else {
      // If no part clients, create a single column for the subclient
      const colLetter = currentCol;
      subClientColumns[subClient.name] = colLetter;

      // Row 5 - Client name
      const cell1 = masterdataSheet.getCell(`${colLetter}5`);
      cell1.value = `${subClient.name.toUpperCase()}`;
      cell1.font = { bold: true, size: 10, name: 'Times New Roman' };
      cell1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell1.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: color }
      };
      cell1.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };

      // Row 6 - DISCOM
      const cell2 = masterdataSheet.getCell(`${colLetter}6`);
      cell2.value = subClient.discom || '';
      cell2.font = { bold: true, size: 10, name: 'Times New Roman' };
      cell2.alignment = { horizontal: 'center', vertical: 'middle' };
      cell2.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: color }
      };
      cell2.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };

      // Row 7 - DIVISION NAME
      const cell3 = masterdataSheet.getCell(`${colLetter}7`);
      masterdataSheet.getRow(7).height = 45;
      cell3.value = subClient.divisionName;
      cell3.font = { bold: true, size: 10, name: 'Times New Roman' };
      cell3.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell3.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: color }
      };
      cell3.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };

      // Row 8 - CONSUMER NO.
      const cell4 = masterdataSheet.getCell(`${colLetter}8`);
      cell4.value = subClient.consumerNo;
      cell4.font = { bold: true, size: 10, name: 'Times New Roman' };
      cell4.alignment = { horizontal: 'center', vertical: 'middle' };
      cell4.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: color }
      };
      cell4.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };

      const cellD9 = masterdataSheet.getCell('D9');
      cellD9.value = '';
      cellD9.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '92D050' } // Bright green in ARGB
      };
      cellD9.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };

      // Row 9 - Blank Row (color-matched to subclient)
      const cellGap = masterdataSheet.getCell(`${colLetter}9`);
      cellGap.value = '';
      cellGap.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: color }
      };
      cellGap.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };

      masterdataSheet.getColumn(colLetter).width = 22;
      currentCol = String.fromCharCode(currentCol.charCodeAt(0) + 1); // Move to next column
    }
  });

  // Add Total DGVCL Share and CHECK-SUM headers after all subclients/partclients
  const totalCol = currentCol;
  const checkSumCol = String.fromCharCode(currentCol.charCodeAt(0) + 1);

  // Total DGVCL Share (merged from row 5 to 8)
  masterdataSheet.mergeCells(`${totalCol}5:${totalCol}8`);
  const totalCell = masterdataSheet.getCell(`${totalCol}5`);
  totalCell.value = 'Total DGVCL Share';
  totalCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  totalCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  totalCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' } // Light gray background
  };
  totalCell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };
  masterdataSheet.getColumn(totalCol).width = 22;

  // Add background color to Total Share column row 9 (blank row)
  const totalCellRow9 = masterdataSheet.getCell(`${totalCol}9`);
  totalCellRow9.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' } // Light gray background
  };

  // CHECK-SUM section - Row 5: "CHECK - SUM with SLDC Approved Data"
  const checkCellRow5 = masterdataSheet.getCell(`${checkSumCol}5`);
  checkCellRow5.value = 'CHECK - SUM with SLDC Approved Data';
  checkCellRow5.font = { bold: true, size: 10, name: 'Times New Roman' };
  checkCellRow5.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  checkCellRow5.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' } // Light gray background
  };
  checkCellRow5.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  // CHECK-SUM section - Rows 6-7-8: "EXCESS INJECTION PPA ***"
  masterdataSheet.mergeCells(`${checkSumCol}6:${checkSumCol}8`);
  const checkCellRow6 = masterdataSheet.getCell(`${checkSumCol}6`);
  checkCellRow6.value = 'EXCESS INJECTION PPA ***';
  checkCellRow6.font = { bold: true, size: 10, name: 'Times New Roman' };
  checkCellRow6.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  checkCellRow6.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFFF' }
  };
  checkCellRow6.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  // Row 9: MWH sub-header only (KWH removed)
  const mwhCell = masterdataSheet.getCell(`${checkSumCol}9`);
  mwhCell.value = 'MWH';
  mwhCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  mwhCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  mwhCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFFF' }
  };
  mwhCell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  masterdataSheet.getColumn(checkSumCol).width = 15;

  // Date range stays fixed at H3:I3 only, does not extend with subclients
  // No need to extend the merge - keep it at H3:I3

  // Data rows (GROSS INJECTION, GROSS DRAWL, NET INJECTION)
  const dataRows = [
    {
      label: 'GROSS INJECTION (MWh)',
      getValues: (data) => {
        const values = [data.mainClient.grossInjectionMWH];

        // Add values for subclients/partclients
        data.subClient.forEach(subClient => {
          const partClients = subClient.subClientsData?.partclient || [];
          if (partClients.length > 0) {
            partClients.forEach(partClient => {
              values.push(partClient.grossInjectionMWHAfterLosses);
            });
          } else {
            values.push(subClient.subClientsData.grossInjectionMWHAfterLosses);
          }
        });

        // Add sum of all subclients/partclients
        const sum = data.subClient.reduce((total, subClient) => {
          const partClients = subClient.subClientsData?.partclient || [];
          if (partClients.length > 0) {
            return total + partClients.reduce((partTotal, partClient) =>
              partTotal + partClient.grossInjectionMWHAfterLosses, 0);
          }
          return total + subClient.subClientsData.grossInjectionMWHAfterLosses;
        }, 0);
        values.push(sum);

        // Calculate CHECK-SUM: Main client value - Sum of all subclients
        const mainClientValue = data.mainClient.grossInjectionMWH;
        const checkSumMWH = mainClientValue - sum;

        // Add CHECK-SUM (MWH only, KWH removed from UI)
        values.push(checkSumMWH); // MWH

        return values;
      },
      format: value => value.toFixed(3)
    },
    {
      label: 'GROSS DRAWL (MWh)',
      getValues: (data) => {
        const values = [data.mainClient.drawlMWH];

        data.subClient.forEach(subClient => {
          const partClients = subClient.subClientsData?.partclient || [];
          if (partClients.length > 0) {
            partClients.forEach(partClient => {
              values.push(partClient.drawlMWHAfterLosses);
            });
          } else {
            values.push(subClient.subClientsData.drawlMWHAfterLosses);
          }
        });

        const sum = data.subClient.reduce((total, subClient) => {
          const partClients = subClient.subClientsData?.partclient || [];
          if (partClients.length > 0) {
            return total + partClients.reduce((partTotal, partClient) =>
              partTotal + partClient.drawlMWHAfterLosses, 0);
          }
          return total + subClient.subClientsData.drawlMWHAfterLosses;
        }, 0);
        values.push(sum);

        // Calculate CHECK-SUM: Main client value - Sum of all subclients
        const mainClientValue = data.mainClient.drawlMWH;
        const checkSumMWH = mainClientValue - sum;

        // Add CHECK-SUM (MWH only, KWH removed from UI)
        values.push(checkSumMWH); // MWH
        return values;
      },
      format: value => value.toFixed(3)
    },
    {
      label: 'NET INJECTION (MWh)',
      getValues: (data) => {
        const values = [data.mainClient.grossInjectionMWH + data.mainClient.drawlMWH];

        data.subClient.forEach(subClient => {
          const partClients = subClient.subClientsData?.partclient || [];
          if (partClients.length > 0) {
            partClients.forEach(partClient => {
              values.push(partClient.netInjectionMWHAfterLosses);
            });
          } else {
            values.push(subClient.subClientsData.netInjectionMWHAfterLosses);
          }
        });

        const sum = data.subClient.reduce((total, subClient) => {
          const partClients = subClient.subClientsData?.partclient || [];
          if (partClients.length > 0) {
            return total + partClients.reduce((partTotal, partClient) =>
              partTotal + partClient.netInjectionMWHAfterLosses, 0);
          }
          return total + subClient.subClientsData.netInjectionMWHAfterLosses;
        }, 0);
        values.push(sum);

        // Calculate CHECK-SUM: Main client value - Sum of all subclients
        const mainClientValue = data.mainClient.grossInjectionMWH + data.mainClient.drawlMWH;
        const checkSumMWH = mainClientValue - sum;

        // Add CHECK-SUM (MWH only, KWH removed from UI)
        values.push(checkSumMWH); // MWH
        return values;
      },
      format: value => value.toFixed(3)
    }
  ];

  // Track CHECK-SUM values to determine header background colors
  const checkSumValues = { mwh: [] };

  dataRows.forEach((row, rowIndex) => {
    const dataRow = 10 + rowIndex;
    const values = row.getValues(lossesCalculationData);

    // Merge A to C for the label
    masterdataSheet.mergeCells(`A${dataRow}:C${dataRow}`);
    const labelCell = masterdataSheet.getCell(`A${dataRow}`);
    masterdataSheet.getRow(dataRow).height = 25;
    labelCell.value = row.label;
    labelCell.font = { bold: true, size: 10, name: 'Times New Roman' };
    labelCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    labelCell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };

    // Value cells start from column D
    // Structure: values[0] = mainClient (D), values[1..n-3] = subclients (E, F, ...), values[n-2] = Total (totalCol), values[n-1] = MWH (checkSumCol)

    // Count total number of data columns (excluding Total and CHECK-SUM)
    // This is: 1 (mainClient) + number of subclients/partclients
    const numDataColumns = values.length - 2; // Exclude Total and MWH (KWH removed)

    values.forEach((value, colIndex) => {
      let colLetter;

      if (colIndex === values.length - 2) {
        // Total column - use the predefined totalCol
        colLetter = totalCol;
      } else if (colIndex === values.length - 1) {
        // CHECK-SUM MWH column
        colLetter = checkSumCol;
      } else {
        // Regular columns: mainClient (colIndex 0) → D, subclients/partclients (colIndex 1 to numDataColumns-1) → E, F, G, ...
        // Calculate column letter: D (68) + colIndex
        colLetter = String.fromCharCode(68 + colIndex); // D = 68, E = 69, F = 70, etc.
      }

      const cell = masterdataSheet.getCell(`${colLetter}${dataRow}`);

      // Apply color to value cells based on column
      const isTotal = colIndex === values.length - 2;
      const isCheckSumMWH = colIndex === values.length - 1;

      // For rows 10-11-12, use default format (3 decimals) for all columns including CHECK-SUM
      // For rows 14+, CHECK-SUM formatting is handled separately in the time block section
      cell.value = typeof value === 'number' ? row.format(value) : value;

      cell.font = { size: 10, bold: true, name: 'Times New Roman' };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };

      // Apply yellow background to CHECK-SUM cells if non-zero
      if (isCheckSumMWH) {
        // Get the original numeric value (before formatting)
        const numericValue = typeof value === 'number' ? value : parseFloat(value) || 0;
        // Check if value is non-zero (with tolerance for floating point errors)
        if (Math.abs(numericValue) > 0.0001) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFF00' } // Yellow background
          };
          // Track the value for header coloring (store original numeric value)
          checkSumValues.mwh.push(numericValue);
        } else {
          // Apply light gray background (#f2f2f2) if zero
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'F2F2F2' } // Light gray background
          };
        }
      }

      // Apply background color to Total Share column (rows 10-12)
      if (isTotal) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'F2F2F2' } // Light gray background
        };
      }

      if (colLetter === 'D') {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '92D050' } // Green for main client          
        };
      } else if (!isTotal && !isCheckSumMWH && colIndex >= 1 && colIndex < values.length - 2) {
        // Find which subclient/partclient this column belongs to
        let color;
        let clientIndex = 0;
        let partClientCount = 0;

        for (let i = 0; i < lossesCalculationData.subClient.length; i++) {
          const subClient = lossesCalculationData.subClient[i];
          const partClients = subClient.subClientsData?.partclient || [];

          if (partClients.length > 0) {
            if (colIndex - 1 >= partClientCount && colIndex - 1 < partClientCount + partClients.length) {
              color = clientColors[i % clientColors.length];
              break;
            }
            partClientCount += partClients.length;
          } else {
            if (colIndex - 1 === partClientCount) {
              color = clientColors[i % clientColors.length];
              break;
            }
            partClientCount += 1;
          }
        }

        if (color) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: color }
          };
        }
      }
    });
  });

  // Apply yellow background to headers if any CHECK-SUM values are non-zero
  const hasNonZeroMWH = checkSumValues.mwh.some(val => Math.abs(val) > 0.0001);

  // Apply yellow background to "EXCESS INJECTION PPA ***" header (rows 6-8) if any non-zero values
  // Otherwise apply light gray background (#f2f2f2)
  if (hasNonZeroMWH) {
    checkCellRow6.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF00' } // Yellow background
    };
  } else {
    checkCellRow6.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'F2F2F2' } // Light gray background
    };
  }

  // Apply yellow background to MWH header (row 9) if any MWH values are non-zero
  // Otherwise apply light gray background (#f2f2f2)
  if (hasNonZeroMWH) {
    mwhCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF00' } // Yellow background
    };
  } else {
    mwhCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'F2F2F2' } // Light gray background
    };
  }

  // Add background color and border to Total Share column row 13
  const totalCellRow13 = masterdataSheet.getCell(`${totalCol}13`);
  totalCellRow13.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' } // Light gray background
  };
  totalCellRow13.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  // Add background color and border to CHECK-SUM column row 13 (MWH only, KWH removed)
  const checkSumMWHCellRow13 = masterdataSheet.getCell(`${checkSumCol}13`);
  checkSumMWHCellRow13.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' } // Light gray background
  };
  checkSumMWHCellRow13.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  // ===== DETAILED TIME BLOCK DATA =====
  const timeBlockStartRow = 13;

  // Time block headers
  const timeBlockHeaders = [
    { cell: `A${timeBlockStartRow}`, value: 'Date', width: 16, bgColor: 'FFFFFF' },
    { cell: `B${timeBlockStartRow}`, value: 'Block Time', width: 13, bgColor: 'FFFFFF' },
    { cell: `C${timeBlockStartRow}`, value: 'Block No', width: 13, bgColor: 'FFFFFF' },
    { cell: `D${timeBlockStartRow}`, value: lossesCalculationData.mainClient.meterNumber, width: 25, bgColor: '92D050' }
  ];

  // Add dynamic headers for each subclient's/partclient's meter with their respective colors
  let timeBlockCol = 'E';
  lossesCalculationData.subClient.forEach((subClient, subIndex) => {
    const color = clientColors[subIndex % clientColors.length];
    const partClients = subClient.subClientsData?.partclient || [];

    if (partClients.length > 0) {
      // Add columns for each part client
      partClients.forEach(partClient => {
        timeBlockHeaders.push({
          cell: `${timeBlockCol}${timeBlockStartRow}`,
          value: `${subClient.meterNumber}`,
          width: 25,
          bgColor: color
        });
        timeBlockCol = String.fromCharCode(timeBlockCol.charCodeAt(0) + 1);
      });
    } else {
      // Add single column for subclient
      timeBlockHeaders.push({
        cell: `${timeBlockCol}${timeBlockStartRow}`,
        value: subClient.meterNumber,
        width: 25,
        bgColor: color
      });
      timeBlockCol = String.fromCharCode(timeBlockCol.charCodeAt(0) + 1);
    }
  });

  // Set column widths and headers for time block data
  timeBlockHeaders.forEach(header => {
    if (header.width) {
      masterdataSheet.getColumn(header.cell.charAt(0)).width = header.width;
    }
    const cell = masterdataSheet.getCell(header.cell);
    cell.value = header.value;
    cell.font = { bold: true, size: 11, name: 'Times New Roman' };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: header.bgColor }
    };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  });

  // Add time block data from the database
  if (lossesCalculationData.mainClient.mainClientMeterDetails &&
    lossesCalculationData.mainClient.mainClientMeterDetails.length > 0) {
    // Get all unique dates from all clients (main and subclients)
    // const timeHeaderRow = 14;
    // Replace the existing date processing code with this:
    const allDates = new Set();

    // Add all possible dates (1st to last day of month)
    for (let day = 1; day <= lastDay; day++) {
      const dateStr = `${day.toString().padStart(2, '0')}-${monthStr}-${lossesCalculationData.year}`;
      allDates.add(dateStr);
    }

    // Convert to array and sort
    const sortedDates = Array.from(allDates).sort();

    // Process all dates (modified section)
    let rowIndex = timeBlockStartRow + 1;

    sortedDates.forEach(date => {
      // Create entries for all 96 blocks (00:00 to 23:45 in 15-minute intervals)
      for (let block = 0; block < 96; block++) {
        const hours = Math.floor(block / 4);
        const minutes = (block % 4) * 15;
        const time = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
        const blockNumber = block + 1;

        // Set date, time and block number with font size 10 (original styling)
        const dateCell = masterdataSheet.getCell(`A${rowIndex}`);
        dateCell.value = date;
        dateCell.font = { size: 10, name: 'Times New Roman' };

        const timeCell = masterdataSheet.getCell(`B${rowIndex}`);
        timeCell.value = time;
        timeCell.font = { size: 10, name: 'Times New Roman' };

        const blockCell = masterdataSheet.getCell(`C${rowIndex}`);
        blockCell.value = blockNumber;
        blockCell.font = { size: 10, name: 'Times New Roman' };

        // Find main client entry for this date/time - use 0 if not found
        const mainEntry = lossesCalculationData.mainClient.mainClientMeterDetails?.find(
          e => e.date === date && e.time === time
        ) || { grossInjectedUnitsTotal: 0 };

        // Main client data (original styling)
        const mainCell = masterdataSheet.getCell(`D${rowIndex}`);
        const mainValue = mainEntry.grossInjectedUnitsTotal;
        mainCell.value = mainValue;
        mainCell.numFmt = '0.00000';
        mainCell.font = { size: 10, name: 'Times New Roman' };

        // Highlight negative or zero values (original styling)
        if (mainValue <= 0) {
          mainCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC7CE' } };
          mainCell.font = { size: 10, color: { argb: '9C0006' }, name: 'Times New Roman' };
        }

        // Sub client data (use 0 if no data available) - original structure
        let subClientsSum = 0;
        let currentCol = 'E'; // Start from column E

        lossesCalculationData.subClient.forEach((subClient, subIndex) => {
          const subEntry = subClient.subClientsData.subClientMeterData?.find(
            e => e.date === date && e.time === time
          ) || { netTotalAfterLosses: 0 };

          const partClients = subClient.subClientsData?.partclient || [];

          if (partClients.length > 0) {
            // Handle part clients - original structure
            partClients.forEach(partClient => {
              const sharingPct = partClient.sharingPercentage / 100;
              const colLetter = currentCol;
              const cell = masterdataSheet.getCell(`${colLetter}${rowIndex}`);

              const partValue = subEntry.netTotalAfterLosses * sharingPct;
              cell.value = partValue;
              cell.numFmt = '0.00000';
              cell.font = { size: 10, name: 'Times New Roman' };
              subClientsSum += partValue;

              if (partValue <= 0) {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC7CE' } };
                cell.font = { size: 10, color: { argb: '9C0006' }, name: 'Times New Roman' };
              }

              currentCol = String.fromCharCode(currentCol.charCodeAt(0) + 1);
            });
          } else {
            // Handle regular subclient - original structure
            const colLetter = currentCol;
            const cell = masterdataSheet.getCell(`${colLetter}${rowIndex}`);

            const value = subEntry.netTotalAfterLosses;
            cell.value = value;
            cell.numFmt = '0.00000';
            cell.font = { size: 10, name: 'Times New Roman' };
            subClientsSum += value;

            if (value <= 0) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC7CE' } };
              cell.font = { size: 10, color: { argb: '9C0006' }, name: 'Times New Roman' };
            }

            currentCol = String.fromCharCode(currentCol.charCodeAt(0) + 1);
          }
        });

        // Calculate CHECK-SUM: Main client value - Sum of all subclients
        const checkSumMWH = mainValue - subClientsSum;

        // CHECK-SUM MWH column - 4 decimal places (KWH removed from UI)
        const checkSumMWHCell = masterdataSheet.getCell(`${checkSumCol}${rowIndex}`);
        checkSumMWHCell.value = checkSumMWH;
        checkSumMWHCell.numFmt = '0.00';
        checkSumMWHCell.font = { size: 10, name: 'Times New Roman' };
        checkSumMWHCell.alignment = { horizontal: 'center', vertical: 'middle' };
        // No border for CHECK-SUM column starting from row 14

        // Style the row with font size 10 (original styling)
        for (let col = 1; col <= timeBlockHeaders.length; col++) {
          const cell = masterdataSheet.getCell(rowIndex, col);
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };
        }
        rowIndex++;
      }
    });
  }
  // ===== ADD thin BORDER AROUND MAIN DATA SECTION (ROWS 5-12) =====
  const lastDataCol = checkSumCol; // This is the CHECK-SUM column

  // Apply thin border around the entire section
  for (let row = 5; row <= 12; row++) {
    for (let col = 1; col <= lastDataCol.charCodeAt(0) - 64; col++) {
      const colLetter = String.fromCharCode(64 + col);
      const cell = masterdataSheet.getCell(`${colLetter}${row}`);

      // Special handling for COMBINED (D), Total (totalCol), and CHECK-SUM (checkSumCol) columns
      const isSpecialColumn = colLetter === 'D' || colLetter === totalCol || colLetter === checkSumCol;

      // Determine border style based on position
      const borderStyles = {
        top: row === 5 || isSpecialColumn ? 'thin' : 'thin',
        bottom: row === 12 ? 'thin' : 'thin',
        left: col === 1 ? 'thin' : 'thin',
        right: colLetter === lastDataCol ? 'thin' : 'thin'
      };

      // For special columns, ensure top border is thin from rows 5-12
      if (isSpecialColumn && row >= 5 && row <= 12) {
        borderStyles.top = 'thin';
      }

      cell.border = {
        top: { style: borderStyles.top, color: { argb: '000000' } },
        left: { style: borderStyles.left, color: { argb: '000000' } },
        bottom: { style: borderStyles.bottom, color: { argb: '000000' } },
        right: { style: borderStyles.right, color: { argb: '000000' } }
      };
    }
  }

  // Special case for the blank row (row 9) - ensure it has thin borders on all sides
  for (let col = 1; col <= lastDataCol.charCodeAt(0) - 64; col++) {
    const colLetter = String.fromCharCode(64 + col);
    const cell = masterdataSheet.getCell(`${colLetter}9`);

    cell.border = {
      top: { style: 'thin', color: { argb: '000000' } },
      left: { style: 'thin', color: { argb: '000000' } },
      bottom: { style: 'thin', color: { argb: '000000' } },
      right: { style: 'thin', color: { argb: '000000' } }
    };
  }

  // Explicitly set thin top borders for special columns (rows 5-12)
  ['D', totalCol, checkSumCol].forEach(colLetter => {
    for (let row = 5; row <= 12; row++) {
      const cell = masterdataSheet.getCell(`${colLetter}${row}`);
      cell.border = {
        ...cell.border,
        top: { style: 'thin', color: { argb: '000000' } }
      };
    }
  });
  // Create the main worksheet for losses calculation
  const worksheet = workbook.addWorksheet('Losses Calculation Sheet');
  // For the Losses Calculation Sheet
  worksheet.pageSetup = worksheetSetup;

  worksheet.properties.tabColor = {
    argb: 'B4C6E7' // This is green color in ARGB format (Alpha, Red, Green, Blue)
  };

  worksheet.pageSetup.orientation = 'landscape'; // Set page orientation to landscape
  worksheet.pageSetup.fitToPage = true; // Fit to page
  worksheet.pageSetup.fitToHeight = 1; // Fit to height
  worksheet.pageSetup.fitToWidth = 1; // Fit to width

  // Add blank row at the top
  worksheet.insertRow(1);

  // Set column widths
  const firstTableColumns = [
    { width: 13 },      // A - Sr. No.
    { width: 21 },     // B - HT Consumer Name
    { width: 17 },     // C - Gross Injected Units
    { width: 17 },     // D - Overall Gross Injected Units
    { width: 17 },     // E - Gross Drawl Units
    { width: 17 },     // F - Overall Gross Drawl Units
    { width: 17 },     // G - Gross Received Units at S/S
    { width: 17 },     // H - Net Drawl Units from S/S
    { width: 17 },     // I - Difference in Injected Units
    { width: 17 },     // J - Difference in Drawl Units
    { width: 17 },     // K - % Weightage According to Gross Injecting Units
    { width: 17 },     // L - % Weightage According to Gross Drawl Units
    { width: 12 },     // M - Losses in Injected Units
    { width: 12 },     // N - in %
    { width: 12 },     // O - Losses in Drawl Units
    { width: 12 },     // P - in %
  ];

  // ===== TITLE ROW =====
  // Title row with merged cells
  worksheet.mergeCells('A2:I2');
  worksheet.mergeCells('J2:L2');
  worksheet.mergeCells('M2:P2');

  const titleCell = worksheet.getCell('A2');
  worksheet.getRow(2).height = 31; // Set height for header row
  titleCell.value = 'Overall Losses & Weightage Calculation Sheet';
  titleCell.font = { bold: true, size: 20, name: 'Times New Roman' };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  titleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '92D050' } // Green background
  };
  titleCell.border = {
    top: { style: 'thin', color: { argb: '000000' } },
    left: { style: 'thin', color: { argb: '000000' } },
    bottom: { style: 'thin', color: { argb: '000000' } },
    right: { style: 'thin', color: { argb: '000000' } }
  };

  const dateCell1 = worksheet.getCell('J2');

  dateCell1.value = `${monthName}-${lossesCalculationData.year.toString().slice(-2)}`;
  dateCell1.font = { bold: true, color: { argb: 'FF0000' }, size: 20, name: 'Times New Roman' }; // Red text
  dateCell1.alignment = { horizontal: 'center', vertical: 'middle' };
  dateCell1.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF00' } // Yellow background
  };
  dateCell1.border = {
    top: { style: 'thin', color: { argb: '000000' } },
    left: { style: 'thin', color: { argb: '000000' } },
    bottom: { style: 'thin', color: { argb: '000000' } },
    right: { style: 'thin', color: { argb: '000000' } }
  };


  const dateCell2 = worksheet.getCell('M2');
  dateCell2.value = `01-${month}-${lossesCalculationData.year} to ${lastDay}-${month}-${lossesCalculationData.year}`;
  dateCell2.font = { bold: true, color: { argb: 'FF0000' }, size: 20, name: 'Times New Roman' };
  dateCell2.alignment = { horizontal: 'center', vertical: 'middle' };
  dateCell2.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF00' } // Yellow background
  };
  dateCell2.border = {
    top: { style: 'thin', color: { argb: '000000' } },
    left: { style: 'thin', color: { argb: '000000' } },
    bottom: { style: 'thin', color: { argb: '000000' } },
    right: { style: 'thin', color: { argb: '000000' } }
  };

  worksheet.insertRow(3);
  worksheet.getRow(3).height = 15; // Set height for header row

  // ===== HEADER ROW =====
  // Header row with merged cells
  // Merge cells first
  worksheet.mergeCells('A4:B4');
  worksheet.mergeCells('C4:F4');
  worksheet.mergeCells('G4:H4');
  worksheet.mergeCells('I4:J4');
  worksheet.mergeCells('K4:L4');
  worksheet.mergeCells('M4:P4');
  worksheet.getRow(4).height = 29;

  // Define headers with their merged ranges
  const headers = [
    { cell: 'A4', value: 'Project Details' },
    { cell: 'C4', value: 'Plant End Generation Details' },
    { cell: 'G4', value: 'S/S End Generation Details' },
    { cell: 'I4', value: 'Difference' },
    { cell: 'K4', value: '% Weightage' },
    { cell: 'M4', value: '% Losses Distribution' }
  ];

  // Apply formatting to each header
  headers.forEach(header => {
    const cell = worksheet.getCell(header.cell);
    cell.value = header.value;
    cell.font = { bold: true, size: 12, name: 'Times New Roman' };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'D9D9D9' } // Light gray background
    };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  });

  // ===== SUB-HEADER ROW =====
  // Sub-header row
  const subHeaders = [
    { cell: 'A5', value: 'Sr. No.' },
    { cell: 'B5', value: 'HT Consumer Name' },
    { cell: 'C5', value: 'Gross Injected Units' },
    { cell: 'D5', value: 'Overall Gross Injected Units' },
    { cell: 'E5', value: 'Gross Drawl Units' },
    { cell: 'F5', value: 'Overall Gross Drawl Units' },
    { cell: 'G5', value: 'Gross Received Units at S/S' },
    { cell: 'H5', value: 'Net Drawl Units from S/S' },
    { cell: 'I5', value: 'Difference in Injected Units, Plant End to S/S End' },
    { cell: 'J5', value: 'Difference in Drawl Units, S/S End to Plant End' },
    { cell: 'K5', value: '% Weightage According to Gross Injecting Units' },
    { cell: 'L5', value: '% Weightage According to Gross Drawl Units' },
    { cell: 'M5', value: 'Losses in Injected Units' },
    { cell: 'N5', value: 'in %' },
    { cell: 'O5', value: 'Losses in Drawl Units' },
    { cell: 'P5', value: 'in %' }
  ];

  subHeaders.forEach(header => {
    const cell = worksheet.getCell(header.cell);
    cell.value = header.value;
    cell.font = { bold: true, size: 11, name: 'Times New Roman' };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF' } // Light gray background
    };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  });

  // Adjust row height for headers
  worksheet.getRow(5).height = 70; // Set height for header row

  // ===== DATA ROWS =====
  // Process sub client data for the table
  let overallGrossInjectedUnits = lossesCalculationData.subClientoverall.overallGrossInjectedUnits;
  let overallGrossDrawlUnits = lossesCalculationData.subClientoverall.grossDrawlUnits;
  let grossReceivedUnits = lossesCalculationData.mainClient.grossInjectionMWH;
  let netDrawlUnits = lossesCalculationData.mainClient.drawlMWH;
  let diffInjectedUnits = lossesCalculationData.difference.diffInjectedUnits;
  let diffDrawlUnits = lossesCalculationData.difference.diffDrawlUnits;


  // Prepare data rows - we'll process both subclients and their partclients
  const dataRowsLosses = [];

  // First, collect all rows we need to display (either subclients or their partclients)
  lossesCalculationData.subClient.forEach((subClient, index) => {
    const subClientData = subClient.subClientsData;

    // Check if this subclient has partclients
    if (subClientData.partclient && subClientData.partclient.length > 0) {
      // Add each partclient as a separate row
      subClientData.partclient.forEach(partClient => {
        dataRowsLosses.push({
          type: 'partclient',
          parentSubClient: subClient,
          data: partClient,
          sharingPercentage: partClient.sharingPercentage
        });
      });
    } else {
      // Add the subclient itself as a row
      dataRowsLosses.push({
        type: 'subclient',
        data: subClient,
        subClientData: subClientData
      });
    }
  });

  // Merged cells for overall values
  if (dataRowsLosses.length > 0) {
    const lastDataRow = 5 + dataRowsLosses.length;

    worksheet.mergeCells(`D6:D${lastDataRow}`);
    worksheet.mergeCells(`F6:F${lastDataRow}`);
    worksheet.mergeCells(`G6:G${lastDataRow}`);
    worksheet.mergeCells(`H6:H${lastDataRow}`);
    worksheet.mergeCells(`I6:I${lastDataRow}`);
    worksheet.mergeCells(`J6:J${lastDataRow}`);

    // Set the merged cell values
    worksheet.getCell('D6').value = overallGrossInjectedUnits.toFixed(3);
    worksheet.getCell('F6').value = overallGrossDrawlUnits.toFixed(3);
    worksheet.getCell('G6').value = grossReceivedUnits.toFixed(3);
    worksheet.getCell('H6').value = netDrawlUnits.toFixed(3);
    worksheet.getCell('I6').value = diffInjectedUnits.toFixed(3);
    worksheet.getCell('J6').value = diffDrawlUnits.toFixed(3);

    // Style the merged cells
    const mergedCells = ['D4', 'F4', 'G4', 'H4', 'I4', 'J4'];
    mergedCells.forEach(cellRef => {
      const cell = worksheet.getCell(cellRef);
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.font = { bold: true, size: 12, name: 'Times New Roman' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'D9D9D9' } // Light gray background
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    // Add data rows for each sub client
    dataRowsLosses.forEach((rowData, index) => {
      const rowIndex = index + 6; // Starting from row 6
      const row = worksheet.getRow(rowIndex);
      row.height = 60; // Set row height

      let consumerName, grossInjection, drawl, weightageInjecting, weightageDrawl;
      let lossesInjectedUnits, inPercentageOfLossesInjectedUnits;
      let lossesDrawlUnits, inPercentageOfLossesDrawlUnits;

      if (rowData.type === 'partclient') {
        // This is a partclient row
        const partClient = rowData.data;
        const parentSubClient = rowData.parentSubClient;

        // Get the index from the partclients array
        const partClients = parentSubClient.subClientsData.partclient;
        const unitIndex = partClients.indexOf(partClient) + 1; // +1 to start from 1 instead of 0

        consumerName = `${parentSubClient.name} - Unit-${unitIndex} (${rowData.sharingPercentage}% of Sharing OA)`;
        grossInjection = partClient.grossInjectionMWH;
        drawl = partClient.drawlMWH;
        weightageInjecting = partClient.weightageGrossInjecting;
        weightageDrawl = partClient.weightageGrossDrawl;
        lossesInjectedUnits = partClient.lossesInjectedUnits;
        inPercentageOfLossesInjectedUnits = partClient.inPercentageOfLossesInjectedUnits;
        lossesDrawlUnits = partClient.lossesDrawlUnits;
        inPercentageOfLossesDrawlUnits = partClient.inPercentageOfLossesDrawlUnits;
      } else {
        const subClient = rowData.data;
        const subClientData = rowData.subClientData;

        consumerName = subClient.name;
        grossInjection = subClientData.grossInjectionMWH;
        drawl = subClientData.drawlMWH;
        weightageInjecting = subClientData.weightageGrossInjecting || 0;
        weightageDrawl = subClientData.weightageGrossDrawl || 0;
        lossesInjectedUnits = subClientData.lossesInjectedUnits || 0;
        inPercentageOfLossesInjectedUnits = subClientData.inPercentageOfLossesInjectedUnits || 0;
        lossesDrawlUnits = subClientData.lossesDrawlUnits || 0;
        inPercentageOfLossesDrawlUnits = subClientData.inPercentageOfLossesDrawlUnits || 0;
      }

      // Set values for each cell
      worksheet.getCell(`A${rowIndex}`).value = index + 1;
      worksheet.getCell(`B${rowIndex}`).value = consumerName;
      worksheet.getCell(`C${rowIndex}`).value = grossInjection.toFixed(3);
      worksheet.getCell(`E${rowIndex}`).value = drawl.toFixed(3);

      worksheet.getCell(`K${rowIndex}`).value = `${weightageInjecting.toFixed(2)}%`;
      worksheet.getCell(`L${rowIndex}`).value = `${weightageDrawl.toFixed(2)}%`;
      worksheet.getCell(`M${rowIndex}`).value = lossesInjectedUnits.toFixed(3);
      worksheet.getCell(`N${rowIndex}`).value = `${inPercentageOfLossesInjectedUnits.toFixed(2)}%`;
      worksheet.getCell(`O${rowIndex}`).value = lossesDrawlUnits.toFixed(3);
      worksheet.getCell(`P${rowIndex}`).value = `${inPercentageOfLossesDrawlUnits.toFixed(2)}%`;

      // Apply formatting to all cells in the row
      for (let col = 1; col <= 16; col++) {
        const cell = worksheet.getCell(rowIndex, col);

        // Default alignment (center)
        let alignment = {
          horizontal: 'center',
          vertical: 'middle',
          wrapText: true
        };

        // Override alignment for column B (Consumer Name)
        if (col === 2) { // Column B (index 2)
          alignment = {
            horizontal: 'left',
            vertical: 'middle',
            wrapText: true
          };
        }

        // Apply formatting
        cell.alignment = alignment;
        cell.font = {
          size: 10,
          name: 'Times New Roman'
        };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFF' } // White background
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      }
    });

    // ===== TOTAL ROW =====
    // Add Total row
    const totalRowIndex = lastDataRow + 1;
    worksheet.getRow(totalRowIndex).height = 34; // Set height for total row
    const totalRow = worksheet.getRow(totalRowIndex);
    totalRow.getCell(1).value = 'Total';
    totalRow.getCell(7).value = grossReceivedUnits.toFixed(3);
    totalRow.getCell(8).value = netDrawlUnits.toFixed(3);
    totalRow.getCell(9).value = diffInjectedUnits.toFixed(3);
    totalRow.getCell(10).value = diffDrawlUnits.toFixed(3);
    totalRow.getCell(11).value = '100%';
    totalRow.getCell(12).value = '100%';
    totalRow.getCell(13).value = diffInjectedUnits.toFixed(3);
    totalRow.getCell(15).value = diffDrawlUnits.toFixed(3);

    // Style the total row
    for (let col = 1; col <= 16; col++) {
      const cell = totalRow.getCell(col);
      cell.font = { bold: true, size: 11, name: 'Times New Roman' };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'D6D6D6' } // Force white background
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
    // ===== ADD thin BORDER AROUND ENTIRE TABLE =====
    const firstDataRow = 4;
    const lastDataRows = 5 + dataRowsLosses.length + 1; // Includes header and total row
    const lastCol = 'P';

    // Apply thin border around the entire table
    for (let row = firstDataRow; row <= lastDataRows; row++) {
      for (let col = 1; col <= lastCol.charCodeAt(0) - 64; col++) {
        const colLetter = String.fromCharCode(64 + col);
        const cell = worksheet.getCell(`${colLetter}${row}`);

        // Determine border style based on position
        const borderStyles = {
          top: row === firstDataRow ? 'thin' : 'thin',
          bottom: row === lastDataRows ? 'thin' : 'thin',
          left: col === 1 ? 'thin' : 'thin',
          right: colLetter === lastCol ? 'thin' : 'thin'
        };

        cell.border = {
          top: { style: borderStyles.top, color: { argb: '000000' } },
          left: { style: borderStyles.left, color: { argb: '000000' } },
          bottom: { style: borderStyles.bottom, color: { argb: '000000' } },
          right: { style: borderStyles.right, color: { argb: '000000' } }
        };
      }
    }

    // Special handling for merged header cells (row 4)
    ['A4:B4', 'C4:F4', 'G4:H4', 'I4:J4', 'K4:L4', 'M4:P4'].forEach(mergeRange => {
      const cell = worksheet.getCell(mergeRange.split(':')[0]);
      cell.border = {
        ...cell.border,
        top: { style: 'thin', color: { argb: '000000' } },
        left: { style: 'thin', color: { argb: '000000' } },
        bottom: { style: 'thin', color: { argb: '000000' } },
        right: { style: 'thin', color: { argb: '000000' } }
      };
    });

    // Special handling for total row to ensure all borders are thin
    for (let col = 1; col <= lastCol.charCodeAt(0) - 64; col++) {
      const colLetter = String.fromCharCode(64 + col);
      const cell = worksheet.getCell(`${colLetter}${lastDataRows}`);

      cell.border = {
        top: { style: 'thin', color: { argb: '000000' } },
        left: { style: col === 1 ? 'thin' : 'thin', color: { argb: '000000' } },
        bottom: { style: 'thin', color: { argb: '000000' } },
        right: { style: colLetter === lastCol ? 'thin' : 'thin', color: { argb: '000000' } }
      };
    }

    // ===== UPDATE SPECIFIC COLUMN BORDERS FOR SUBCLIENT DATA ROWS =====
    const firstSubclientDataRow = 6;
    const lastSubclientDataRow = 5 + dataRowsLosses.length; // Last row of subclient data

    // Columns that need thin left border (Gross Injected Units - C)
    const columnsWiththinLeftBorder = ['C'];

    // Columns that need thin right border (Overall Gross Drawl - F, Net Drawl - H, 
    // Difference in Drawl - J, % Weightage Drawl - L)
    const columnsWiththinRightBorder = ['F', 'H', 'J', 'L'];

    // Apply thin borders to specified columns for all data rows
    for (let currentRow = firstSubclientDataRow; currentRow <= lastSubclientDataRow; currentRow++) {
      // thin left borders
      columnsWiththinLeftBorder.forEach(columnLetter => {
        const targetCell = worksheet.getCell(`${columnLetter}${currentRow}`);
        targetCell.border = {
          ...targetCell.border,
          left: { style: 'thin', color: { argb: '000000' } }
        };
      });

      // thin right borders
      columnsWiththinRightBorder.forEach(columnLetter => {
        const targetCell = worksheet.getCell(`${columnLetter}${currentRow}`);
        targetCell.border = {
          ...targetCell.border,
          right: { style: 'thin', color: { argb: '000000' } }
        };
      });
    }

    // Also update the header row (row 5) for these columns
    columnsWiththinLeftBorder.forEach(columnLetter => {
      const headerCell = worksheet.getCell(`${columnLetter}5`);
      headerCell.border = {
        ...headerCell.border,
        left: { style: 'thin', color: { argb: '000000' } }
      };
    });

    columnsWiththinRightBorder.forEach(columnLetter => {
      const headerCell = worksheet.getCell(`${columnLetter}5`);
      headerCell.border = {
        ...headerCell.border,
        right: { style: 'thin', color: { argb: '000000' } }
      };
    });

    // Update the total row borders for these columns
    const totalRowNumber = 5 + dataRowsLosses.length + 1;
    columnsWiththinLeftBorder.forEach(columnLetter => {
      const totalRowCell = worksheet.getCell(`${columnLetter}${totalRowNumber}`);
      totalRowCell.border = {
        ...totalRowCell.border,
        left: { style: 'thin', color: { argb: '000000' } }
      };
    });

    columnsWiththinRightBorder.forEach(columnLetter => {
      const totalRowCell = worksheet.getCell(`${columnLetter}${totalRowNumber}`);
      totalRowCell.border = {
        ...totalRowCell.border,
        right: { style: 'thin', color: { argb: '000000' } }
      };
    });
    // Merge the Total cell across columns
    worksheet.mergeCells(`A${totalRowIndex}:B${totalRowIndex}`);

    // ===== ABT METER RAW DATA SECTION =====
    // Add space after the first table
    const abtStartRow = totalRowIndex + 2;

    const baseABTColumns = [
      { width: 15 },      // Q - Date (assuming this continues from first table)
      { width: 15 },      // R - Block Time
      { width: 15 },      // S - Block No
      { width: 15 }       // T - SLDC APPROVED DATA
    ];

    // Calculate dynamic columns for ABT section
    const subClientColumns = [];
    subClients.forEach(subClient => {
      const subClientData = subClient.subClientsData;
      if (subClientData.partclient && subClientData.partclient.length > 0) {
        subClientColumns.push({ width: 15 }); // Total column
        subClientData.partclient.forEach(() => {
          subClientColumns.push({ width: 15 }); // Partclient columns
        });
      } else {
        subClientColumns.push({ width: 15 }, { width: 15 });
      }
    });

    worksheet.columns = [...firstTableColumns, ...baseABTColumns, ...subClientColumns];

    worksheet.getRow(abtStartRow).height = 95; // Fixed height for header row

    // ABT METER RAW DATA header
    worksheet.mergeCells(`A${abtStartRow}:C${abtStartRow}`);
    const abtHeaderCell = worksheet.getCell(`A${abtStartRow}`);
    abtHeaderCell.value = "ABT METER RAW DATA";
    abtHeaderCell.font = { bold: true, size: 11, name: 'Times New Roman' };
    abtHeaderCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    abtHeaderCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF00" }
    };
    abtHeaderCell.border = {
      top: { style: "thin", color: { argb: '000000' } },
      left: { style: "thin", color: { argb: '000000' } },
      bottom: { style: "thin", color: { argb: '000000' } },
      right: { style: "thin", color: { argb: '000000' } }
    };

    // SLDC APPROVED DATA header
    const sldcHeaderCell = worksheet.getCell(`D${abtStartRow}`);
    sldcHeaderCell.value = "SLDC APPROVED DATA";
    sldcHeaderCell.font = { bold: true, size: 12, name: 'Times New Roman' };
    sldcHeaderCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    sldcHeaderCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "92D050" }
    };
    sldcHeaderCell.border = {
      top: { style: "thin", color: { argb: '000000' } },
      left: { style: "thin", color: { argb: '000000' } },
      bottom: { style: "thin", color: { argb: '000000' } },
      right: { style: "thin", color: { argb: '000000' } }
    };

    // Client colors array extended to support up to 10 subclients
    const clientColors = [
      'FFC000', // Orange
      'B4C6E7', // Light blue
      'E6B9D8', // Light purple
      'F8CBAD', // Peach
      'C6E0B4', // Light green
      'D9D9D9', // Light gray
      'E2EFDA', // Very light green
      'BDD7EE', // Light blue
      'FFF2CC', // Light yellow
      'DDEBF7'  // Very light blue
    ];

    // Create headers for each sub client with different colors
    subClients.forEach((subClient, index) => {
      const subClientData = subClient.subClientsData;
      const hasPartClients = subClientData.partclient && subClientData.partclient.length > 0;
      const columnsNeeded = hasPartClients ? 1 + subClientData.partclient.length : 2;

      const startCol = 5 + (index > 0 ?
        subClients.slice(0, index).reduce((sum, sc) => {
          const scHasPart = sc.subClientsData.partclient && sc.subClientsData.partclient.length > 0;
          return sum + (scHasPart ? 1 + sc.subClientsData.partclient.length : 2);
        }, 0) : 0);

      const endCol = startCol + columnsNeeded - 1;
      const startLetter = worksheet.getColumn(startCol).letter;
      const endLetter = worksheet.getColumn(endCol).letter;
      const range = `${startLetter}${abtStartRow}:${endLetter}${abtStartRow}`;

      // Unmerge if already merged (safety check)
      if (worksheet.getCell(`${startLetter}${abtStartRow}`).isMerged) {
        worksheet.unMergeCells(range);
      }

      // Merge cells
      worksheet.mergeCells(range);

      const clientHeaderCell = worksheet.getCell(`${startLetter}${abtStartRow}`);
      clientHeaderCell.value = subClient.name;
      clientHeaderCell.font = { bold: true, size: 12, name: 'Times New Roman' };
      clientHeaderCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      clientHeaderCell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: clientColors[index % clientColors.length] },
      };
      clientHeaderCell.border = {
        top: { style: "thin", color: { argb: '000000' } },
        left: { style: "thin", color: { argb: '000000' } },
        bottom: { style: "thin", color: { argb: '000000' } },
        right: { style: "thin", color: { argb: '000000' } }
      };

      // Set column widths for the additional columns
      for (let col = startCol; col <= endCol; col++) {
        worksheet.getColumn(col).width = 15;
      }
    });

    const totalRow2 = abtStartRow + 1;

    // Clear previous values (optional but helps avoid merge issues)
    ["A", "B", "C"].forEach((col) => {
      worksheet.getCell(`${col}${totalRow2}`).value = null;
      worksheet.getCell(`${col}${totalRow2 + 1}`).value = null;
    });

    // Unmerge if already merged to prevent merge error
    try {
      worksheet.unMergeCells(`A${totalRow2}:C${totalRow2 + 1}`);
    } catch (err) {
      // Ignore if not merged
    }

    // Now merge ABC cells across 2 rows
    worksheet.mergeCells(`A${totalRow2}:C${totalRow2 + 1}`);

    const summaryHeaderCell = worksheet.getCell(`A${totalRow2}`);
    summaryHeaderCell.value = ""; // Or "GROSS INJECTION in MWH" if needed
    summaryHeaderCell.font = { bold: true, size: 12, name: 'Times New Roman' };
    summaryHeaderCell.alignment = { horizontal: "center", vertical: "middle" };
    summaryHeaderCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF00" },
    };
    summaryHeaderCell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };

    // D column - merge and style
    worksheet.mergeCells(`D${totalRow2}:D${totalRow2 + 1}`);
    const totalColDCell = worksheet.getCell(`D${totalRow2}`);
    totalColDCell.value = "Total";
    totalColDCell.font = { bold: true, size: 11, name: 'Times New Roman' };
    totalColDCell.alignment = { horizontal: "center", vertical: "middle" };
    totalColDCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "92D050" },
    };
    totalColDCell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };

    // Clear the cell below (just in case)
    // worksheet.getCell(`D${totalRow2 + 1}`).value = "";

    // Track the current column position
    let currentCol = 5;

    // Add "Total" and "NET Total after Losses" headers for subclients
    subClients.forEach((subClient, index) => {
      const subClientData = subClient.subClientsData;
      const color = clientColors[index % clientColors.length];
      const hasPartClients = subClientData.partclient && subClientData.partclient.length > 0;

      if (hasPartClients) {
        // Total column for parent subclient (merged across two rows)
        const totalColLetter = worksheet.getColumn(currentCol).letter;
        worksheet.mergeCells(`${totalColLetter}${totalRow2}:${totalColLetter}${totalRow2 + 1}`);

        const totalCell = worksheet.getCell(`${totalColLetter}${totalRow2}`);
        totalCell.value = "Total";
        totalCell.font = { bold: true, name: 'Times New Roman' };
        totalCell.alignment = { horizontal: "center", vertical: "middle" };
        totalCell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: color },
        };
        totalCell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };

        // Move to the next column for partclients
        currentCol++;

        // Columns for each partclient
        subClientData.partclient.forEach((partClient, partIndex) => {
          const partCol = currentCol + partIndex;
          const partColLetter = worksheet.getColumn(partCol).letter;

          worksheet.getColumn(partCol).width = 15;

          // Top row: UNIT-X (XX% Sharing)
          const unitCell = worksheet.getCell(`${partColLetter}${totalRow2}`);
          unitCell.value = `${partClient.divisionName} \n (${partClient.sharingPercentage}% Sharing)`;
          unitCell.font = { bold: true, name: 'Times New Roman' };
          unitCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
          unitCell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: color },
          };
          unitCell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };

          // Second row: NET Total after Losses
          const netCell = worksheet.getCell(`${partColLetter}${totalRow2 + 1}`);
          netCell.value = "NET Total after Losses";
          netCell.font = { bold: true, name: 'Times New Roman' };
          netCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
          netCell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: color },
          };
          netCell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        });

        // Update current column position after all partclients
        currentCol += subClientData.partclient.length;
      } else {
        // No partclients - use same layout but with different content
        const colLetter = worksheet.getColumn(currentCol).letter;

        // Merge both rows for this subclient's "Total"
        worksheet.mergeCells(`${colLetter}${totalRow2}:${colLetter}${totalRow2 + 1}`);
        const totalCell = worksheet.getCell(`${colLetter}${totalRow2}`);
        totalCell.value = "Total";
        totalCell.font = { bold: true, name: 'Times New Roman' };
        totalCell.alignment = { horizontal: "center", vertical: "middle" };
        totalCell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: color },
        };
        totalCell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };

        // Move to next column for "NET Total after Losses"
        currentCol++;
        const netColLetter = worksheet.getColumn(currentCol).letter;
        worksheet.mergeCells(`${netColLetter}${totalRow2}:${netColLetter}${totalRow2 + 1}`);
        const netCell = worksheet.getCell(`${netColLetter}${totalRow2}`);
        netCell.value = "NET Total after Losses";
        netCell.font = { bold: true, name: 'Times New Roman' };
        netCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        netCell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: color },
        };
        netCell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };

        // Update current column
        currentCol++;
      }
    });

    // Adjust the data rows below to account for the additional rows if partclients exist
    const hasPartClients = subClients.some(subClient =>
      subClient.subClientsData.partclient && subClient.subClientsData.partclient.length > 0
    );

    // Add GROSS INJECTION, DRAWL, and NET INJECTION rows
    const grossInjectionRow = totalRow2 + (hasPartClients ? 2 : 2);
    const drawlRow = grossInjectionRow + 1;
    const netInjectionRow = drawlRow + 1;

    // First unmerge any existing merged cells in these ranges to prevent errors
    ['A', 'B', 'C'].forEach(col => {
      [grossInjectionRow, drawlRow, netInjectionRow].forEach(row => {
        const cellAddress = `${col}${row}`;
        if (worksheet.getCell(cellAddress).isMerged) {
          const mergedCell = worksheet.getCell(cellAddress);
          worksheet.unMergeCells(mergedCell.master.address);
        }
      });
    });


    // Set values and merge cells
    worksheet.getCell(`A${grossInjectionRow}`).value = "GROSS INJECTION in MWH";
    worksheet.getCell(`A${drawlRow}`).value = "DRAWL in MWH";
    worksheet.getCell(`A${netInjectionRow}`).value = "NET INJECTION in MWH";

    // Merge ABC columns for each row
    [grossInjectionRow, drawlRow, netInjectionRow].forEach(row => {
      const range = `A${row}:C${row}`;
      if (!worksheet.getCell(`A${row}`).isMerged) {
        worksheet.mergeCells(range);
      }
    });

    // Style the merged cells
    [grossInjectionRow, drawlRow, netInjectionRow].forEach((row) => {
      const cell = worksheet.getCell(`A${row}`);
      cell.font = { bold: true, size: 10, name: 'Times New Roman' };
      cell.alignment = { horizontal: "left", vertical: "middle" };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFF00" }, // Yellow background
      };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });

    // Add values for SLDC APPROVED DATA (now in column D)
    worksheet.getCell(`D${grossInjectionRow}`).value = grossReceivedUnits.toFixed(3);
    worksheet.getCell(`D${grossInjectionRow}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "92D050" },
    };
    worksheet.getCell(`D${drawlRow}`).value = netDrawlUnits.toFixed(3);
    worksheet.getCell(`D${drawlRow}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "92D050" },
    };
    worksheet.getCell(`D${netInjectionRow}`).value = (grossReceivedUnits + netDrawlUnits).toFixed(3);
    worksheet.getCell(`D${netInjectionRow}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "92D050" },
    };

    let dataStartCol = 5;

    // Add values for each sub client
    subClients.forEach((subClient, index) => {
      const subClientData = subClient.subClientsData;
      const color = clientColors[index % clientColors.length];

      if (subClientData.partclient && subClientData.partclient.length > 0) {
        // Handle subclient with partclients
        const totalCol = dataStartCol;
        const totalColLetter = worksheet.getColumn(totalCol).letter;

        // Values under TOTAL
        [grossInjectionRow, drawlRow, netInjectionRow].forEach(row => {
          const cell = worksheet.getCell(`${totalColLetter}${row}`);
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };
        });

        worksheet.getCell(`${totalColLetter}${grossInjectionRow}`).value = subClientData.grossInjectionMWH.toFixed(3);
        worksheet.getCell(`${totalColLetter}${drawlRow}`).value = subClientData.drawlMWH.toFixed(3);
        worksheet.getCell(`${totalColLetter}${netInjectionRow}`).value = subClientData.netInjectionMWH.toFixed(3);

        // Values under each partclient (NET Total after Losses)
        subClientData.partclient.forEach((partClient, i) => {
          const partCol = dataStartCol + 1 + i;
          const partColLetter = worksheet.getColumn(partCol).letter;

          const grossAfterLosses = partClient.grossInjectionMWHAfterLosses || partClient.grossInjectionMWH - partClient.lossesInjectedUnits;
          const drawlAfterLosses = partClient.drawlMWHAfterLosses || partClient.drawlMWH - partClient.lossesDrawlUnits;
          const netAfterLosses = grossAfterLosses + drawlAfterLosses;

          [grossInjectionRow, drawlRow, netInjectionRow].forEach(row => {
            const cell = worksheet.getCell(`${partColLetter}${row}`);
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };
          });

          worksheet.getCell(`${partColLetter}${grossInjectionRow}`).value = grossAfterLosses.toFixed(3);
          worksheet.getCell(`${partColLetter}${drawlRow}`).value = drawlAfterLosses.toFixed(3);
          worksheet.getCell(`${partColLetter}${netInjectionRow}`).value = netAfterLosses.toFixed(3);
        });

        dataStartCol += 1 + subClientData.partclient.length;
      } else {
        // Handle subclient without partclients
        const totalCol = dataStartCol;
        const netCol = dataStartCol + 1;
        const totalColLetter = worksheet.getColumn(totalCol).letter;
        const netColLetter = worksheet.getColumn(netCol).letter;

        // TOTAL values
        [grossInjectionRow, drawlRow, netInjectionRow].forEach(row => {
          const cell = worksheet.getCell(`${totalColLetter}${row}`);
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };
        });

        worksheet.getCell(`${totalColLetter}${grossInjectionRow}`).value = subClientData.grossInjectionMWH.toFixed(3);
        worksheet.getCell(`${totalColLetter}${drawlRow}`).value = subClientData.drawlMWH.toFixed(3);
        worksheet.getCell(`${totalColLetter}${netInjectionRow}`).value = subClientData.netInjectionMWH.toFixed(3);

        // NET values
        const grossAfterLosses = subClientData.grossInjectionMWHAfterLosses || subClientData.grossInjectionMWH - subClientData.lossesInjectedUnits;
        const drawlAfterLosses = subClientData.drawlMWHAfterLosses || subClientData.drawlMWH - subClientData.lossesDrawlUnits;
        const netAfterLosses = grossAfterLosses + drawlAfterLosses;

        [grossInjectionRow, drawlRow, netInjectionRow].forEach(row => {
          const cell = worksheet.getCell(`${netColLetter}${row}`);
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };
        });

        worksheet.getCell(`${netColLetter}${grossInjectionRow}`).value = grossAfterLosses.toFixed(3);
        worksheet.getCell(`${netColLetter}${drawlRow}`).value = drawlAfterLosses.toFixed(3);
        worksheet.getCell(`${netColLetter}${netInjectionRow}`).value = netAfterLosses.toFixed(3);

        dataStartCol += 2;
      }
    });


    // Style all value cells in these rows
    for (let row = grossInjectionRow; row <= netInjectionRow; row++) {
      for (let col = 4; col <= 13; col++) { // Columns D to M
        const cell = worksheet.getCell(row, col);
        if (cell.value !== undefined && cell.value !== null) {
          cell.alignment = { horizontal: "center", vertical: "middle" };
          cell.font = { size: 10, bold: true, name: 'Times New Roman' };
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        }
      }
    }

    // Add Date, Block Time, Block No headers
    const timeHeaderRow = netInjectionRow + 1;
    worksheet.getCell(`A${timeHeaderRow}`).value = "Date";
    worksheet.getCell(`A${timeHeaderRow}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF00" }, // Yellow (matches ABT METER RAW DATA header)
    };
    worksheet.getCell(`A${timeHeaderRow}`).alignment = { horizontal: "center", vertical: "middle" };
    worksheet.getCell(`A${timeHeaderRow}`).font = { size: 11, bold: true, name: 'Times New Roman' };
    worksheet.getCell(`B${timeHeaderRow}`).value = "Block Time";
    worksheet.getCell(`B${timeHeaderRow}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF00" }, // Yellow
    };
    worksheet.getCell(`B${timeHeaderRow}`).alignment = { horizontal: "center", vertical: "middle" };
    worksheet.getCell(`B${timeHeaderRow}`).font = { size: 11, bold: true, name: 'Times New Roman' };
    worksheet.getCell(`C${timeHeaderRow}`).value = "Block No";
    worksheet.getCell(`C${timeHeaderRow}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF00" }, // Yellow
    };

    worksheet.getCell(`C${timeHeaderRow}`).alignment = { horizontal: "center", vertical: "middle" };
    worksheet.getCell(`C${timeHeaderRow}`).font = { size: 11, bold: true, name: 'Times New Roman' };

    // Add meter numbers for each client (SLDC in column D)
    // SLDC meter number (column D)
    worksheet.getCell(`D${timeHeaderRow}`).value = lossesCalculationData.mainClient.meterNumber;
    worksheet.getCell(`D${timeHeaderRow}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "92D050" }, // Green (matches SLDC header)
    };
    worksheet.getCell(`D${timeHeaderRow}`).alignment = { horizontal: "center", vertical: "middle" };
    worksheet.getCell(`D${timeHeaderRow}`).font = { size: 11, bold: true, name: 'Times New Roman' };

    let meterStartCol = 5; // Starting column for meter numbers

    subClients.forEach((subClient, index) => {
      const subClientData = subClient.subClientsData;
      const color = clientColors[index % clientColors.length];
      const meterNumber = subClient.meterNumber;

      if (subClientData.partclient && subClientData.partclient.length > 0) {
        // Subclient with partclients
        const totalCols = 1 + subClientData.partclient.length;

        for (let i = 0; i < totalCols; i++) {
          const col = meterStartCol + i;
          const cell = worksheet.getCell(worksheet.getColumn(col).letter + timeHeaderRow);


          // Only first column gets the meter number
          if (i === 0) {
            cell.value = meterNumber;
            cell.alignment = { horizontal: "center", vertical: "middle" };
            cell.font = { size: 11, bold: true, name: 'Times New Roman' };
            cell.border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: color },
          };
          cell.alignment = { horizontal: "center", vertical: "middle" };
          cell.font = { size: 11, bold: true, name: 'Times New Roman' };
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        }

        meterStartCol += totalCols;
      } else {
        // Subclient without partclients — 2 columns (TOTAL + NET)
        for (let i = 0; i < 2; i++) {
          const col = meterStartCol + i;
          const cell = worksheet.getCell(worksheet.getColumn(col).letter + timeHeaderRow);

          if (i === 0) {
            cell.value = meterNumber;
            cell.alignment = { horizontal: "center", vertical: "middle" };
            cell.font = { size: 11, bold: true, name: 'Times New Roman' };
            cell.border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: color },
          };
          cell.alignment = { horizontal: "center", vertical: "middle" };
          cell.font = { size: 11, bold: true, name: 'Times New Roman' };
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        }

        meterStartCol += 2;
      }
    });

    // Add actual time blocks data from the database
    if (lossesCalculationData.mainClient.mainClientMeterDetails &&
      lossesCalculationData.mainClient.mainClientMeterDetails.length > 0) {
      // Replace the existing date processing code with this:
      let lastColumn = 4; // Start with column D (SLDC data)

      lossesCalculationData.subClient.forEach(subClient => {
        const subClientData = subClient.subClientsData;
        if (subClientData.partclient && subClientData.partclient.length > 0) {
          // For subclients with partclients: 1 column for total + 1 column per partclient
          lastColumn += 1 + subClientData.partclient.length;
        } else {
          // For regular subclients: 2 columns (TOTAL + NET)
          lastColumn += 2;
        }
      });

      const lastColumnLetter = worksheet.getColumn(lastColumn).letter;

      // Now continue with the date processing code
      const allDates = new Set();

      // Add all possible dates (1st to last day of month)
      for (let day = 1; day <= lastDay; day++) {
        const dateStr = `${day.toString().padStart(2, '0')}-${monthStr}-${lossesCalculationData.year}`;
        allDates.add(dateStr);
      }

      // Convert to array and sort
      const sortedDates = Array.from(allDates).sort();

      // Add time blocks for all dates
      let rowIndex = timeHeaderRow + 1;

      sortedDates.forEach(date => {
        // Create entries for all 96 blocks (00:00 to 23:45 in 15-minute intervals)
        for (let block = 0; block < 96; block++) {
          const hours = Math.floor(block / 4);
          const minutes = (block % 4) * 15;
          const time = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
          const blockNumber = block + 1;

          // Set date, time and block number
          worksheet.getCell(`A${rowIndex}`).value = date;
          worksheet.getCell(`B${rowIndex}`).value = time;
          worksheet.getCell(`C${rowIndex}`).value = blockNumber;

          // Find main client entry for this date/time - use 0 if not found
          const mainEntry = lossesCalculationData.mainClient.mainClientMeterDetails?.find(
            e => e.date === date && e.time === time
          ) || { grossInjectedUnitsTotal: 0 };

          // Add main client data (column D)
          const mainGrossCell = worksheet.getCell(`D${rowIndex}`);
          const mainGrossValue = mainEntry.grossInjectedUnitsTotal;
          mainGrossCell.value = mainGrossValue;
          mainGrossCell.numFmt = '0.000';

          if (mainGrossValue <= 0) {
            mainGrossCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC7CE" } };
            mainGrossCell.font = { color: { argb: "9C0006" } };
          }

          // Add sub client data
          let currentCol = 5;
          lossesCalculationData.subClient.forEach((subClient, clientIndex) => {
            const subClientData = subClient.subClientsData;

            // Find subclient entry for this date/time - use 0 if not found
            const subEntry = subClientData.subClientMeterData?.find(
              e => e.date === date && e.time === time
            ) || { grossInjectedUnitsTotal: 0, netTotalAfterLosses: 0 };

            if (subClientData.partclient && subClientData.partclient.length > 0) {
              // Handle subclient with partclients
              // Total column
              const grossCell = worksheet.getCell(`${String.fromCharCode(64 + currentCol)}${rowIndex}`);
              const grossValue = subEntry.grossInjectedUnitsTotal;
              grossCell.value = grossValue;
              grossCell.numFmt = '0.000';

              if (grossValue <= 0) {
                grossCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC7CE" } };
                grossCell.font = { color: { argb: "9C0006" } };
              }

              // Partclient columns
              subClientData.partclient.forEach((partClient, partIndex) => {
                const partCol = currentCol + 1 + partIndex;
                const netLossCell = worksheet.getCell(`${String.fromCharCode(64 + partCol)}${rowIndex}`);

                // Calculate net total after losses
                const sharingPercentage = partClient.sharingPercentage;
                const netValue = (subEntry.netTotalAfterLosses * sharingPercentage) / 100;

                netLossCell.value = netValue;
                netLossCell.numFmt = '0.000';

                if (netValue <= 0) {
                  netLossCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC7CE" } };
                  netLossCell.font = { color: { argb: "9C0006" } };
                }
              });

              currentCol += 1 + subClientData.partclient.length;
            } else {
              // Handle subclient without partclients
              // Total column
              const grossCell = worksheet.getCell(`${String.fromCharCode(64 + currentCol)}${rowIndex}`);
              const grossValue = subEntry.grossInjectedUnitsTotal;
              grossCell.value = grossValue;
              grossCell.numFmt = '0.000';

              if (grossValue <= 0) {
                grossCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC7CE" } };
                grossCell.font = { color: { argb: "9C0006" } };
              }

              // NET column
              const netLossCell = worksheet.getCell(`${String.fromCharCode(64 + currentCol + 1)}${rowIndex}`);
              const netValue = subEntry.netTotalAfterLosses;
              netLossCell.value = netValue;
              netLossCell.numFmt = '0.000';

              if (netValue <= 0) {
                netLossCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC7CE" } };
                netLossCell.font = { color: { argb: "9C0006" } };
              }

              currentCol += 2;
            }
          });

          // Style the row
          for (let col = 1; col <= lastColumn; col++) {
            const cell = worksheet.getCell(rowIndex, col);
            cell.alignment = { horizontal: "center", vertical: "middle" };
            cell.border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }

          rowIndex++;
        }
      });

      // Apply borders to all header rows dynamically
      const headerRows = [abtStartRow, totalRow2, grossInjectionRow, drawlRow, netInjectionRow, timeHeaderRow];
      headerRows.forEach(row => {
        for (let col = 1; col <= lastColumn; col++) {
          const cell = worksheet.getCell(row, col);
          if (cell.value !== undefined) {
            cell.border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }
        }
      });
    }
  }


  return workbook;
};

const downloadLossesCalculationExcel = async (req, res) => {
  try {
    const { id } = req.params;

    // Get the losses calculation data from database
    const lossesCalculationData = await LossesCalculationData.findById(id).lean();
    if (!lossesCalculationData) {
      return res.status(404).json({ message: 'Losses calculation data not found' });
    }

    // Generate Excel workbook
    const workbook = await exportLossesCalculationToExcel(lossesCalculationData);

    // Format the date components
    const monthNames = [
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ];
    const monthName = monthNames[lossesCalculationData.month - 1];
    const year = lossesCalculationData.year;
    const formattedDate = `Month of ${monthName}-${year}`;

    // Get the name components
    const mainClientName = lossesCalculationData.mainClient.mainClientDetail?.name?.trim() || '';
    const feederName = lossesCalculationData.mainClient.mainClientDetail?.subTitle?.trim() || '';

    // Sanitize filename components (remove special characters and extra spaces)
    const sanitize = (str) => str.replace(/[^a-zA-Z0-9\s-]/g, '').replace(/\s+/g, ' ').trim();

    // Build filename in required format
    const sanitizedMainClient = sanitize(mainClientName);
    const sanitizedFeeder = sanitize(feederName);
    const fileName = feederName
      ? `${sanitizedMainClient} (${sanitizedFeeder}) ${formattedDate}.xlsx`
      : `${sanitizedMainClient} ${formattedDate}.xlsx`;

    // Set response headers
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${encodeURIComponent(fileName)}"`
    );
    res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');

    // Write the workbook to the response
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    res.status(500).json({ message: 'Error generating Excel file', error: error.message });
  }
};

const getLossesDataLastFourMonths = async (req, res) => {
  try {
    const { mainClientId, month, year } = req.body;

    // Validate input fields
    if (!mainClientId || !month || !year) {
      return res.status(400).json({ message: 'MainClientId, Month, and Year are required.' });
    }

    // Validate month and year formats
    const paddedMonth = month.toString().padStart(2, '0');
    if (!moment(`${year}-${paddedMonth}-01`, 'YYYY-MM-DD', true).isValid()) {
      return res.status(400).json({ message: 'Invalid month or year format.' });
    }

    const startDate = moment(`${year}-${paddedMonth}-01`, 'YYYY-MM-DD').startOf('month');
    const monthsToFetch = [];

    // Collect the last 4 months
    for (let i = 0; i < 4; i++) {
      monthsToFetch.push({
        month: parseInt(startDate.clone().subtract(i, 'months').format('M')),
        year: parseInt(startDate.clone().subtract(i, 'months').format('YYYY')),
      });
    }

    const results = [];

    // Fetch data for the last 4 months
    for (const period of monthsToFetch) {
      try {
        const data = await LossesCalculationData.findOne({
          mainClientId,
          month: period.month,
          year: period.year,
        })
          .select('month year subClient.name subClient.subClientsData.grossInjectionMWHAfterLosses subClient.subClientsData.weightageGrossInjecting subClient.subClientsData.partclient.divisionName subClient.subClientsData.partclient.grossInjectionMWHAfterLosses subClient.subClientsData.partclient.weightageGrossInjecting');

        if (data) {
          const structuredData = {
            month: data.month,
            year: data.year,
            subClients: data.subClient.map(sub => ({
              name: sub.name,
              grossInjectionMWHAfterLosses: sub.subClientsData.grossInjectionMWHAfterLosses,
              weightageGrossInjecting: sub.subClientsData.weightageGrossInjecting,
              partClients: sub.subClientsData.partclient ? sub.subClientsData.partclient.map(part => ({
                divisionName: part.divisionName,
                grossInjectionMWHAfterLosses: part.grossInjectionMWHAfterLosses,
                weightageGrossInjecting: part.weightageGrossInjecting,
              })) : [],
            })),
          };
          results.push(structuredData);
        } else {
          logger.warn(`No data found for Main Client ${mainClientId} for ${period.month}/${period.year}`);
        }
      } catch (error) {
        logger.error(`Error fetching data for ${period.month}/${period.year}: ${error.message}`);
      }
    }

    if (results.length === 0) {
      return res.status(404).json({ message: 'No data found for the requested months.' });
    }

    res.status(200).json({
      message: 'Losses Calculation data fetched successfully.',
      data: results,
    });

  } catch (error) {
    // General error handling for any issues
    logger.error(`Error fetching losses data for MainClientId: ${mainClientId}, Month: ${month}, Year: ${year}: ${error.message}`);
    res.status(500).json({ message: 'Internal server error', error: error.message });
  }
};

// Get SLDC GROSS INJECTION and DRAWL for a specific main client, month, and year
const getSLDCData = async (req, res) => {
  try {
    const { mainClientId, month, year } = req.body;

    // Validate input
    if (!mainClientId || !month || !year) {
      return res.status(400).json({ message: "mainClientId, month, and year are required." });
    }

    // Get the latest report for this main client, month, and year
    const sldcData = await LossesCalculationData.findOne({
      mainClientId,
      month: Number(month),
      year: Number(year)
    })
      .sort({ updatedAt: -1 })   // <-- Add this line to always get the latest!
      .select('SLDCGROSSINJECTION SLDCGROSSDRAWL DGVCL MGVCL PGVCL UGVCL TAECO TSECO TEL');

    if (!sldcData) {
      return res.status(404).json({ message: "No SLDC data found for this client and period." });
    }

    // Build response data - only include DISCOM values if they are not null/undefined and not 0
    const responseData = {
      SLDCGROSSINJECTION: sldcData.SLDCGROSSINJECTION,
      SLDCGROSSDRAWL: sldcData.SLDCGROSSDRAWL
    };

    // Add DISCOM values only if they exist and are not 0
    if (sldcData.DGVCL != null && sldcData.DGVCL !== 0) {
      responseData.DGVCL = sldcData.DGVCL;
    }
    if (sldcData.MGVCL != null && sldcData.MGVCL !== 0) {
      responseData.MGVCL = sldcData.MGVCL;
    }
    if (sldcData.PGVCL != null && sldcData.PGVCL !== 0) {
      responseData.PGVCL = sldcData.PGVCL;
    }
    if (sldcData.UGVCL != null && sldcData.UGVCL !== 0) {
      responseData.UGVCL = sldcData.UGVCL;
    }
    if (sldcData.TAECO != null && sldcData.TAECO !== 0) {
      responseData.TAECO = sldcData.TAECO;
    }
    if (sldcData.TSECO != null && sldcData.TSECO !== 0) {
      responseData.TSECO = sldcData.TSECO;
    }
    if (sldcData.TEL != null && sldcData.TEL !== 0) {
      responseData.TEL = sldcData.TEL;
    }

    res.status(200).json({
      message: "SLDC data fetched successfully.",
      data: responseData
    });
  } catch (error) {
    logger.error(`Error fetching SLDC data: ${error.message}`);
    res.status(500).json({ message: "Error fetching SLDC data.", error: error.message });
  }
};



module.exports = { generateLossesCalculation, getLatestLossesReports, downloadLossesCalculationExcel, getLossesDataLastFourMonths, getSLDCData };
