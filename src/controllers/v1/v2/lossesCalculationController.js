const LossesCalculationData = require('../../models/v1/lossesCalculation.model');
const moment = require('moment');
const MainClient = require('../../models/v1/mainClient.model');
const SubClient = require('../../models/v1/subClient.model');
const MeterData = require('../../models/v1/meterData.model');
const PartClient = require('../../models/v1/partClient.model');
const logger = require('../../utils/logger');  // Importing the logger
const ExcelJS = require('exceljs');

// Controller: generateLossesCalculation
// Mirrors Excel logic end-to-end. Main client 15-min values follow your LET() formula:
//   - Build baseline s = SUMIFS(Load Survey ...)
//   - p = SUMIF(>0 of baseline), n = SUMIF(<0 of baseline)
//   - f = 1 + IF(s<=0, (MAIN_ENTRY!L9 - n)/n, (MAIN_ENTRY!B9 - p)/p)  == (target / sum_by_sign)
//   - main interval = s * f * (-SUMMARY!E12/1000)  -> mapped here as (AE * mf * pn)/1000 * f

const parseOr = (val, fallback, { forbidZero = false } = {}) => {
  const n = Number(val);
  if (!Number.isFinite(n)) return fallback;
  if (forbidZero && n === 0) return fallback;
  return n;
};

// Assumes models + logger available:
// const { MainClient, SubClient, PartClient, MeterData, LossesCalculationData } = require('...');
// const logger = require('...');

const generateLossesCalculation = async (req, res) => {
  try {
    const { mainClientId, month, year } = req.body;

    // SLDC monthly approved (SUMMARY!A8 positive, SUMMARY!K8 negative)
    const SLDC_INJ = req.body.SLDCGROSSINJECTION;  // MAIN ENTRY!B9 equivalent (positive)
    let SLDC_DRW = req.body.SLDCGROSSDRAWL;      // MAIN ENTRY!L9 equivalent (negative)
    if (typeof SLDC_DRW === 'number' && SLDC_DRW > 0) SLDC_DRW = -Math.abs(SLDC_DRW);

    if (!mainClientId || !month || !year) {
      return res.status(400).json({ message: "Missing required parameters: mainClientId, month, year." });
    }

    // STEP 0: cache on SLDC inputs (legacy behavior)
    const existing = await LossesCalculationData.findOne({
      mainClientId, month, year,
      ...(SLDC_INJ !== undefined && { SLDCGROSSINJECTION: SLDC_INJ }),
      ...(SLDC_DRW !== undefined && { SLDCGROSSDRAWL: SLDC_DRW }),
    });
    if (existing) {
      existing.updatedAt = new Date();
      await existing.save();
      return res.status(200).json({ message: "Existing calculation data retrieved successfully.", data: existing });
    }

    // STEP 1: main client
    const mainClientData = await MainClient.findById(mainClientId);
    if (!mainClientData) return res.status(404).json({ message: "Main Client not found" });

    // STEP 2: sub clients
    const subClients = await SubClient.find({ mainClient: mainClientId });
    if (!subClients.length) return res.status(404).json({ message: "No Sub Clients found" });

    // STEP 2.1: part clients (ok if none)
    const partClientsData = {};
    await Promise.all(
      subClients.map(async (sc) => {
        try { partClientsData[sc._id] = await PartClient.find({ subClient: sc._id }); }
        catch { partClientsData[sc._id] = []; }
      })
    );

    // STEP 3: meter data (fallback to check meter)
    const clientsUsingCheckMeter = [];
    const subsUsingCheck = [];

    let mainClientMeterData = await MeterData.find({
      meterNumber: mainClientData.abtMainMeter?.meterNumber, month, year
    });

    if (!mainClientMeterData.length && mainClientData.abtCheckMeter?.meterNumber) {
      mainClientMeterData = await MeterData.find({
        meterNumber: mainClientData.abtCheckMeter.meterNumber, month, year
      });
      if (mainClientMeterData.length) clientsUsingCheckMeter.push(mainClientData.name);
    }

    if (!mainClientMeterData.length) {
      return res.status(400).json({
        message: "Meter data missing for Main Client. Both abtMainMeter and abtCheckMeter files are missing."
      });
    }

    const subClientMeterData = {};
    const missingSubClientMeters = [];
    await Promise.all(
      subClients.map(async (sc) => {
        let data = await MeterData.find({
          meterNumber: sc.abtMainMeter?.meterNumber, month, year
        });

        if (!data.length && sc.abtCheckMeter?.meterNumber) {
          data = await MeterData.find({
            meterNumber: sc.abtCheckMeter.meterNumber, month, year
          });
          if (data.length) subsUsingCheck.push(sc.name);
        }

        if (!data.length) missingSubClientMeters.push(sc.name);
        subClientMeterData[sc._id] = data;
      })
    );

    if (missingSubClientMeters.length) {
      return res.status(400).json({
        message: `Meter data missing for Sub Clients: ${missingSubClientMeters.join(', ')}. Both abtMainMeter and abtCheckMeter files are missing.`
      });
    }

    // STEP 4: init doc
    let doc = new LossesCalculationData({
      mainClientId, month, year,
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
          mf: mainClientData.mf,  // maps to SUMMARY!E12 in your sheet
          sharingPercentage: mainClientData.sharingPercentage,
          contactNo: mainClientData.contactNo,
          email: mainClientData.email
        },
        grossInjectionMWH: 0,
        drawlMWH: 0,
        netInjectionMWH: 0,
        mainClientMeterDetails: []
      },
      subClient: [],
      subClientoverall: { overallGrossInjectedUnits: 0, grossDrawlUnits: 0 },
      difference: { diffInjectedUnits: 0, diffDrawlUnits: 0 }
    });

    // STEP 5: MAIN baseline, then apply per-sign factor f like your LET() formula
    // s_base = (AE * mf * pn)/1000  (this equals s * (-SUMMARY!E12/1000) with pn = -1)
    const mainMF = parseOr(mainClientData.mf, 1, { forbidZero: true });
    const mainPN = parseOr(mainClientData.pn, -1, { forbidZero: true }); // usually -1

    // 5a) First pass: compute baseline and collect p/n
    const baseline = []; // store per-interval baseline before scaling
    let pSum = 0; // Σ positive baseline
    let nSum = 0; // Σ negative baseline
    mainClientMeterData.forEach((meter) => {
      meter.dataEntries.forEach((entry) => {
        const rawAE = entry.parameters['Bidirectional Active(I-E)'] ?? entry.parameters['Net Active'];
        const ae = parseOr(rawAE, NaN);
        if (!Number.isFinite(ae)) return;

        const s_base = (ae * mainMF * mainPN) / 1000; // matches your -SUMMARY!E12/1000 sign/scale
        const time = entry.parameters['Interval Start'];
        const row = {
          date: entry.parameters.Date,
          time,
          s_base
        };
        baseline.push(row);
        if (s_base > 0) pSum += s_base;
        else if (s_base < 0) nSum += s_base;
      });
    });

    // 5b) Compute factors like f = 1 + (target - sum)/sum == target/sum
    const fPos = (pSum !== 0 && typeof SLDC_INJ === 'number') ? (SLDC_INJ / pSum) : 1;
    const fNeg = (nSum !== 0 && typeof SLDC_DRW === 'number') ? (SLDC_DRW / nSum) : 1;

    // 5c) Second pass: write scaled values to doc.mainClient.mainClientMeterDetails
    for (const row of baseline) {
      const scaled =
        row.s_base > 0 ? row.s_base * fPos :
          row.s_base < 0 ? row.s_base * fNeg : 0;

      doc.mainClient.mainClientMeterDetails.push({
        date: row.date,
        time: row.time,
        grossInjectedUnitsTotal: scaled,   // final main per-interval (your HELPER TWO/THREE effect)
        preSLDC: row.s_base,               // keep original baseline for audit if you want
        scaleFactor: row.s_base > 0 ? fPos : (row.s_base < 0 ? fNeg : 1)
      });

      if (scaled > 0) doc.mainClient.grossInjectionMWH += scaled;
      else if (scaled < 0) doc.mainClient.drawlMWH += scaled;
    }

    doc.mainClient.netInjectionMWH =
      doc.mainClient.grossInjectionMWH + doc.mainClient.drawlMWH;

    // === From here on, everything is identical to the last version you approved ===

    // STEP 6: counters for sub totals (raw)
    let overallPosRaw = 0;
    let overallNegRaw = 0;

    // STEP 7: build SUB client raw + redistribution keys
    const mainByKey = new Map(); // date__time -> MAIN (after scale) for redistribution (HELPER!G)
    for (const it of (doc.mainClient.mainClientMeterDetails || [])) {
      const key = `${it.date}__${it.time}`;
      mainByKey.set(key, Number(it.grossInjectedUnitsTotal) || 0);
    }

    for (const sc of subClients) {
      const meterData = subClientMeterData[sc._id];
      if (!meterData || !meterData.length) continue;

      const { meterNumber, meterType } = meterData[0];

      const scData = {
        name: sc.name,
        divisionName: sc.divisionName,
        consumerNo: sc.consumerNo,
        contactNo: sc.contactNo,
        email: sc.email,
        subClientId: sc._id,
        meterNumber, meterType,
        discom: sc.discom,
        voltageLevel: sc.voltageLevel,
        ctptSrNo: sc.ctptSrNo,
        ctRatio: sc.ctRatio,
        ptRatio: sc.ptRatio,
        mf: sc.mf,
        acCapacityKw: sc.acCapacityKw,
        subClientsData: {
          grossInjectionMWH: 0,
          drawlMWH: 0,
          netInjectionMWH: 0,
          subClientMeterData: [],
          // legacy fields
          weightageGrossInjecting: 0,
          weightageGrossDrawl: 0,
          lossesInjectedUnits: 0,
          inPercentageOfLossesInjectedUnits: 0,
          lossesDrawlUnits: 0,
          inPercentageOfLossesDrawlUnits: 0
        }
      };

      meterData.forEach((meter) => {
        meter.dataEntries.forEach((entry) => {
          const v = entry.parameters['Bidirectional Active(I-E)'] ?? entry.parameters['Net Active'];
          if (v === undefined) return;
          const Eraw = (v * sc.mf * sc.pn) / 1000;
          const t = entry.parameters['Interval Start'];

          scData.subClientsData.subClientMeterData.push({
            date: entry.parameters.Date,
            time: t,
            grossInjectedUnitsTotal: Eraw, // raw sub
            redistributedToMain: 0,         // will become E (Losses!E)
            netTotalAfterLosses: 0
          });

          if (Eraw > 0) scData.subClientsData.grossInjectionMWH += Eraw;
          else if (Eraw < 0) scData.subClientsData.drawlMWH += Eraw;
        });
      });

      scData.subClientsData.netInjectionMWH =
        scData.subClientsData.grossInjectionMWH + scData.subClientsData.drawlMWH;

      const parts = partClientsData[sc._id];
      if (parts && parts.length) {
        scData.subClientsData.partclient = parts.map((p) => ({
          divisionName: p.divisionName,
          consumerNo: p.consumerNo,
          sharingPercentage: Number(p.sharingPercentage),
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
          inPercentageOfLossesDrawlUnits: 0
        }));
      }

      doc.subClient.push(scData);

      overallPosRaw += scData.subClientsData.grossInjectionMWH;
      overallNegRaw += scData.subClientsData.drawlMWH;
    }

    doc.subClientoverall.overallGrossInjectedUnits = overallPosRaw;
    doc.subClientoverall.grossDrawlUnits = overallNegRaw;

    // STEP 7.1: proportional re-distribution to MAIN (HELPER!G → Losses!E)
    const totalSubByKey = new Map();
    for (const sc of doc.subClient) {
      for (const it of sc.subClientsData.subClientMeterData) {
        const key = `${it.date}__${it.time}`;
        totalSubByKey.set(key, (totalSubByKey.get(key) || 0) + (Number(it.grossInjectedUnitsTotal) || 0));
      }
    }
    for (const sc of doc.subClient) {
      for (const it of sc.subClientsData.subClientMeterData) {
        const key = `${it.date}__${it.time}`;
        const mainV = mainByKey.get(key) ?? 0;
        const subTot = totalSubByKey.get(key) ?? 0;
        it.redistributedToMain = (subTot !== 0) ? (it.grossInjectedUnitsTotal * (mainV / subTot)) : 0;
      }
    }

    // STEP 8: legacy differences (reporting vs main)
    doc.difference.diffInjectedUnits = doc.subClientoverall.overallGrossInjectedUnits - doc.mainClient.grossInjectionMWH;
    doc.difference.diffDrawlUnits = doc.subClientoverall.grossDrawlUnits - doc.mainClient.drawlMWH;

    // STEP 9: legacy weightages
    for (const sc of doc.subClient) {
      const sd = sc.subClientsData;
      sd.weightageGrossInjecting = overallPosRaw !== 0 ? (sd.grossInjectionMWH / overallPosRaw) * 100 : 0;
      sd.weightageGrossDrawl = overallNegRaw !== 0 ? (sd.drawlMWH / overallNegRaw) * 100 : 0;
      sd.lossesInjectedUnits = (doc.difference.diffInjectedUnits * sd.weightageGrossInjecting) / 100;
      sd.inPercentageOfLossesInjectedUnits = sd.grossInjectionMWH !== 0 ? (sd.lossesInjectedUnits / sd.grossInjectionMWH) * 100 : 0;
      sd.lossesDrawlUnits = (doc.difference.diffDrawlUnits * sd.weightageGrossDrawl) / 100;
      sd.inPercentageOfLossesDrawlUnits = sd.drawlMWH !== 0 ? (sd.lossesDrawlUnits / sd.drawlMWH) * 100 : 0;
    }

    // STEP 10: N6/P6 on redistributed E with residual snap
    let sumPosE = 0, sumNegE = 0;
    for (const sc of doc.subClient) {
      for (const it of sc.subClientsData.subClientMeterData) {
        const E = Number(it.redistributedToMain) || 0;
        if (E > 0) sumPosE += E;
        else if (E < 0) sumNegE += E;
      }
    }

    const N6 = (typeof SLDC_INJ === 'number' && sumPosE !== 0) ? (sumPosE - SLDC_INJ) / sumPosE : 0;
    const P6 = (typeof SLDC_DRW === 'number' && sumNegE !== 0) ? (sumNegE - SLDC_DRW) / sumNegE : 0;
    doc.lossRates = { injectionLossFraction_N6: N6, drawlLossFraction_P6: P6 };

    let firstPosRef = null, firstNegRef = null;

    for (const sc of doc.subClient) {
      sc.subClientsData.grossInjectionMWHAfterLosses = 0;
      sc.subClientsData.drawlMWHAfterLosses = 0;

      for (let idx = 0; idx < sc.subClientsData.subClientMeterData.length; idx++) {
        const it = sc.subClientsData.subClientMeterData[idx];
        if (!it.date) { it.netTotalAfterLosses = 0; continue; }

        const E = Number(it.redistributedToMain) || 0;
        const rate = (E > 0) ? N6 : (E < 0 ? P6 : 0);
        const net = E - rate * E;

        if (E > 0 && firstPosRef === null) firstPosRef = { sc, idx };
        if (E < 0 && firstNegRef === null) firstNegRef = { sc, idx };

        it.netTotalAfterLosses = net;
        if (net > 0) sc.subClientsData.grossInjectionMWHAfterLosses += net;
        else sc.subClientsData.drawlMWHAfterLosses += net;
      }

      sc.subClientsData.netInjectionMWHAfterLosses =
        sc.subClientsData.grossInjectionMWHAfterLosses + sc.subClientsData.drawlMWHAfterLosses;

      // part clients after-loss (safe if none)
      if (sc.subClientsData.partclient && sc.subClientsData.partclient.length > 0) {
        sc.subClientsData.partclient.forEach(p => {
          p.grossInjectionMWHAfterLosses = 0;
          p.drawlMWHAfterLosses = 0;
          p.netInjectionMWHAfterLosses = 0;
        });

        for (const it of sc.subClientsData.subClientMeterData) {
          it.partclient = sc.subClientsData.partclient.map(p => {
            const share = (p.sharingPercentage || 0) / 100;
            const val = (it.netTotalAfterLosses || 0) * share;
            return { divisionName: p.divisionName, netTotalAfterLosses: val };
          });

          it.partclient.forEach((pc, i) => {
            if (pc.netTotalAfterLosses > 0)
              sc.subClientsData.partclient[i].grossInjectionMWHAfterLosses += pc.netTotalAfterLosses;
            else if (pc.netTotalAfterLosses < 0)
              sc.subClientsData.partclient[i].drawlMWHAfterLosses += pc.netTotalAfterLosses;
          });
        }

        sc.subClientsData.partclient.forEach(p => {
          p.netInjectionMWHAfterLosses = p.grossInjectionMWHAfterLosses + p.drawlMWHAfterLosses;
        });
      }
    }

    // Residual snap to match SLDC exactly
    let sumNetPos = 0, sumNetNeg = 0;
    for (const sc of doc.subClient) {
      sumNetPos += Math.max(sc.subClientsData.grossInjectionMWHAfterLosses || 0, 0);
      sumNetNeg += Math.min(sc.subClientsData.drawlMWHAfterLosses || 0, 0);
    }
    const targetPos = (typeof SLDC_INJ === 'number') ? SLDC_INJ : sumNetPos;
    const targetNeg = (typeof SLDC_DRW === 'number') ? SLDC_DRW : sumNetNeg;

    const posResidual = targetPos - sumNetPos;
    const negResidual = targetNeg - sumNetNeg;

    if (Math.abs(posResidual) > 1e-9 && firstPosRef) {
      const { sc, idx } = firstPosRef;
      const it = sc.subClientsData.subClientMeterData[idx];
      it.netTotalAfterLosses += posResidual;
      sc.subClientsData.grossInjectionMWHAfterLosses += posResidual;
      sc.subClientsData.netInjectionMWHAfterLosses =
        sc.subClientsData.grossInjectionMWHAfterLosses + sc.subClientsData.drawlMWHAfterLosses;
    }
    if (Math.abs(negResidual) > 1e-9 && firstNegRef) {
      const { sc, idx } = firstNegRef;
      const it = sc.subClientsData.subClientMeterData[idx];
      it.netTotalAfterLosses += negResidual;
      sc.subClientsData.drawlMWHAfterLosses += negResidual;
      sc.subClientsData.netInjectionMWHAfterLosses =
        sc.subClientsData.grossInjectionMWHAfterLosses + sc.subClientsData.drawlMWHAfterLosses;
    }

    // STEP 12: record SLDC deltas vs MAIN baseline (intervals already correct)
    if (typeof SLDC_INJ === 'number') doc.SLDCGROSSINJECTION = SLDC_INJ;
    if (typeof SLDC_DRW === 'number') doc.SLDCGROSSDRAWL = SLDC_DRW;
    if (typeof SLDC_INJ === 'number' || typeof SLDC_DRW === 'number') {
      doc.mainClient.asperApprovedbySLDCGROSSINJECTION = (typeof SLDC_INJ === 'number') ? (SLDC_INJ - doc.mainClient.grossInjectionMWH) : 0;
      doc.mainClient.asperApprovedbySLDCGROSSDRAWL = (typeof SLDC_DRW === 'number') ? (SLDC_DRW - doc.mainClient.drawlMWH) : 0;
    }

    // STEP 13: save + respond
    await doc.save();

    const allUsingCheck = [...clientsUsingCheckMeter, ...subsUsingCheck];
    if (allUsingCheck.length) doc.clientsUsingCheckMeter = allUsingCheck;

    return res.status(200).json({
      message: 'Losses Calculation successfully completed.',
      data: doc,
      ...(allUsingCheck.length > 0 && { clientsUsingCheckMeter: allUsingCheck })
    });

  } catch (err) {
    logger.error(`Error generating Losses Calculation Data: ${err.message}`);
    return res.status(500).json({ message: 'Error generating Losses Calculation Data', error: err.message });
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
    margins: {
      left: 0.2,
      right: 0.2,
      top: 0.2,
      bottom: 0.2,
      header: 0.2,
      footer: 0.2
    },
    horizontalCentered: true,
    verticalCentered: false,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    paperSize: 9 // A4
  };

  // Create a First sheet for Summary Sheet
  const summarySheet = workbook.addWorksheet('SUMMARY')
  summarySheet.pageSetup = worksheetSetup;

  // Set tab color (using exceljs)
  summarySheet.properties.tabColor = {
    argb: 'FFFF00' // This is green color in ARGB format (Alpha, Red, Green, Blue)
  };

  const month = lossesCalculationData.month < 10 ? `0${lossesCalculationData.month}` : lossesCalculationData.month;
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const monthName = monthNames[lossesCalculationData.month - 1]; // Adjusted for 0-based index

  // Prepare data rows
  const subClients = lossesCalculationData.subClient;

  // Calculate required date variables
  const monthStr = lossesCalculationData.month < 10
    ? `0${lossesCalculationData.month}`
    : lossesCalculationData.month.toString();
  const lastDays = new Date(
    lossesCalculationData.year,
    lossesCalculationData.month,
    0
  ).getDate();

  const lastDay = new Date(lossesCalculationData.year, lossesCalculationData.month, 0).getDate();

  // Set column widths as per the image
  summarySheet.columns = [
    { width: 7 },  // A - Sr. No.
    { width: 30 },  // B - HT Consumer Name
    { width: 13 },  // C - HT Consumer No.
    { width: 22 },  // D - Wheeling Division Office/Location
    { width: 22 },  // E - Wheeling Discom
    { width: 22 },  // F - Project Capacity (kW) (AC)
    { width: 22 },  // G - Share in Gross Injected Units to Panetha S/S (MWh)
    { width: 22 },  // H - Share in Gross Drawl Units from Panetha S/S (MWh)
    { width: 22 },  // I - Net Injected Units to Panetha S/S (MWh)
    { width: 22 }   // J - % Weightage According to Gross Injecting
  ];

  // Helper function to display exact values
  const displayExactValue = (value) => {
    if (value === undefined || value === null || isNaN(value)) return '0.000';

    // Convert to number and format to exactly 3 decimal places
    const numValue = Number(value);

    // Handle cases where rounding might add extra decimals
    const formattedValue = numValue.toLocaleString('en-US', {
      minimumFractionDigits: 3,
      maximumFractionDigits: 3,
      useGrouping: false // Don't add thousands separators
    });

    return formattedValue;
  };

  // Add blank row at the top
  summarySheet.getRow(1).height = 15;

  // Calculate the last column needed based on number of subclients
  const lastColumnForClients = String.fromCharCode(69 + lossesCalculationData.subClient.length); // 69='E', +1 for main client
  const lastColumnForHeader = lastColumnForClients; // Same as last client column

  // Row 2: Title row with SUMMARY SHEET and Company Name
  const titleRow = summarySheet.getRow(2);
  titleRow.height = 42;

  // SUMMARY SHEET cell - merge A2:C2
  summarySheet.mergeCells('A2:C2');
  const summaryTitleCell = summarySheet.getCell('A2');
  summaryTitleCell.value = 'SUMMARY SHEET';
  summaryTitleCell.font = { bold: true, size: 16, name: 'Times New Roman', color: { argb: '000000' } };
  summaryTitleCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  summaryTitleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF00' } // Yellow background
  };
  summaryTitleCell.border = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' }
  };

  // Company name cell - merge D2:J2
  summarySheet.mergeCells('D2:J2');
  const companyCellSummary = summarySheet.getCell('D2');
  const acCapacityMwSummary = (lossesCalculationData.mainClient.mainClientDetail.acCapacityKw / 1000).toFixed(2);
  companyCellSummary.value = `${lossesCalculationData.mainClient.mainClientDetail.name.toUpperCase()} - ${acCapacityMwSummary} MW AC Generation Details`;
  companyCellSummary.font = { bold: true, size: 16, name: 'Times New Roman', color: { argb: '000000' } };
  companyCellSummary.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  companyCellSummary.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF00' } // Yellow background
  };
  companyCellSummary.border = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' }
  };

  // Row 3: Month row
  const monthRow = summarySheet.getRow(3);
  monthRow.height = 30;

  // Month label (merge A3:C3)
  summarySheet.mergeCells('A3:C3');
  const monthLabelCell = summarySheet.getCell('C3');
  monthLabelCell.value = 'Month';
  monthLabelCell.font = { bold: true, size: 14, name: 'Times New Roman', color: { argb: 'FF0000' } };
  monthLabelCell.alignment = { horizontal: 'center', vertical: 'middle' };
  monthLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '92D050' } // Green background
  };
  monthLabelCell.border = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' }
  };

  // Month value (merge D3:J3)
  summarySheet.mergeCells('D3:J3');
  const monthValueCell = summarySheet.getCell('D3');
  monthValueCell.value = `${monthName}-${lossesCalculationData.year.toString().slice(-2)}`;
  monthValueCell.font = { bold: true, size: 18, name: 'Times New Roman', color: { argb: 'FF0000' } };
  monthValueCell.alignment = { horizontal: 'center', vertical: 'middle' };
  monthValueCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '92D050' } // Green background
  };
  monthValueCell.border = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' }
  };

  // Row 4: Generation Period row
  const periodRow = summarySheet.getRow(4);
  periodRow.height = 30;

  // Generation Period label (merge A4:C4 - changed from A4:D4)
  summarySheet.mergeCells('A4:C4');
  const periodLabelCell = summarySheet.getCell('A4'); // Changed from C4 to A4
  periodLabelCell.value = 'Generation Period';
  periodLabelCell.font = { bold: true, size: 14, name: 'Times New Roman', color: { argb: 'FF0000' } };
  periodLabelCell.alignment = { horizontal: 'center', vertical: 'middle' };
  periodLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '92D050' } // Green background
  };
  periodLabelCell.border = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' }
  };

  // Generation Period value (merge D4:J4)
  const periodEndColumn = 'J';
  summarySheet.mergeCells(`D4:${periodEndColumn}4`);
  const periodValueCell = summarySheet.getCell('D4');

  // Format dates as DD-MM-YYYY
  const startDate = `01-${monthStr}-${lossesCalculationData.year}`;
  const endDate = `${lastDay}-${monthStr}-${lossesCalculationData.year}`;

  periodValueCell.value = `${startDate} to ${endDate}`;
  periodValueCell.font = { bold: true, size: 18, name: 'Times New Roman', color: { argb: 'FF0000' } };
  periodValueCell.alignment = { horizontal: 'center', vertical: 'middle' };
  periodValueCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '92D050' } // Green background
  };
  periodValueCell.border = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' }
  };
  // Row 5: CPP CLIENTS header row - merge A5 to last column
  const cppRow = summarySheet.getRow(5);
  cppRow.height = 40;
  summarySheet.mergeCells(`A5:J5`);
  const cppCell = summarySheet.getCell('A5');
  cppCell.value = `CPP CLIENTS - ${lossesCalculationData.mainClient.mainClientDetail.name.toUpperCase()} (Lead generator) SOLAR PLANT WITH INJECTION TO ${lossesCalculationData.mainClient.mainClientDetail.subTitle} AT 11kv, ABT METER: ${lossesCalculationData.mainClient.meterNumber}`;
  cppCell.font = { bold: true, size: 12, name: 'Times New Roman', color: { argb: '0000cc' } };
  cppCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  cppCell.border = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' }
  };

  // Row 6: SLDC APPROVED header row
  const sldcRow = summarySheet.getRow(6);
  sldcRow.height = 70;

  // SLDC APPROVED label - merge A6:B6
  summarySheet.mergeCells('A6:B6');
  const sldcLabelCell = summarySheet.getCell('A6');
  sldcLabelCell.value = 'SLDC APPROVED';
  sldcLabelCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  sldcLabelCell.alignment = { horizontal: 'center', vertical: 'middle' };
  sldcLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '92D050' } // Green background
  };
  sldcLabelCell.border = {
    top: { style: 'thin' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'thin' }
  };

  // Total (MWh) label - C6
  const totalLabelCell = summarySheet.getCell('C6');
  totalLabelCell.value = 'Total (MWh)';
  totalLabelCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  totalLabelCell.alignment = { horizontal: 'center', vertical: 'middle' };
  totalLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '92D050' } // Green background
  };
  totalLabelCell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'medium' },
    right: { style: 'thin' }
  };

  // Feeder Name label - D6
  const feederLabelCell = summarySheet.getCell('D6');
  feederLabelCell.value = 'Feeder Name =>';
  feederLabelCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  feederLabelCell.alignment = { horizontal: 'right', vertical: 'middle' };
  feederLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'D9D9D9' }
  };
  feederLabelCell.border = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' }
  };

  // Main client cell - E6
  const mainClientCell = summarySheet.getCell('E6');
  mainClientCell.value = `(Lead Generator)\n${lossesCalculationData.mainClient.mainClientDetail.name.toUpperCase()}`;
  mainClientCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  mainClientCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  mainClientCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'D9D9D9' }
  };
  mainClientCell.border = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' }
  };

  // Create 5 subclient cells (F6 to J6)
  for (let i = 0; i < 5; i++) {
    const col = String.fromCharCode(70 + i); // 70 is 'F'
    const cellRef = summarySheet.getCell(`${col}6`);

    // If there's a subclient at this index, use its name, otherwise empty string
    cellRef.value = lossesCalculationData.subClient[i]
      ? lossesCalculationData.subClient[i].name.toUpperCase()
      : '';

    cellRef.font = { bold: true, size: 10, name: 'Times New Roman' };
    cellRef.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cellRef.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'D9D9D9' }
    };
    cellRef.border = {
      top: { style: 'medium' },
      left: { style: 'medium' },
      bottom: { style: 'medium' },
      right: { style: 'medium' }
    };
  }

  // Calculate totals - using exact values from database
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

  // Use SLDC values if available, otherwise use calculated values
  const grossInjectedValue = lossesCalculationData.SLDCGROSSINJECTION !== undefined
    ? lossesCalculationData.SLDCGROSSINJECTION
    : mainClientGrossInjection;

  const grossDrawlValue = lossesCalculationData.SLDCGROSSDRAWL !== undefined
    ? lossesCalculationData.SLDCGROSSDRAWL
    : mainClientDrawl;

  const netInjectedValue = grossInjectedValue + grossDrawlValue;

  // Row 7: Gross Injected Units row
  const grossInjectedRow = summarySheet.getRow(7);
  grossInjectedRow.height = 45;

  // Gross Injected Units label - merge A7:B7
  summarySheet.mergeCells('A7:B7');
  const grossInjectedLabelCell = summarySheet.getCell('A7');
  grossInjectedLabelCell.value = `Gross Injected Units to ${lossesCalculationData.mainClient.mainClientDetail.subTitle}`;
  grossInjectedLabelCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  grossInjectedLabelCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
  grossInjectedLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '92D050' } // Green background
  };
  grossInjectedLabelCell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  // Gross Injected Units value - C7
  const grossInjectedValueCell = summarySheet.getCell('C7');
  grossInjectedValueCell.value = displayExactValue(grossInjectedValue);
  grossInjectedValueCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  grossInjectedValueCell.alignment = { horizontal: 'center', vertical: 'middle' };
  grossInjectedValueCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '92D050' } // Green background
  };
  grossInjectedValueCell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  // ABT Main Meter label - D7
  const abtMeterLabelCell = summarySheet.getCell('D7');
  abtMeterLabelCell.value = 'ABT Main Meter Sr. No.';
  abtMeterLabelCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  abtMeterLabelCell.alignment = { horizontal: 'right', vertical: 'middle', wrapText: true };
  abtMeterLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'D9D9D9' }
  };
  abtMeterLabelCell.border = {
    top: { style: 'thin' },
    left: { style: 'medium' },
    bottom: { style: 'thin' },
    right: { style: 'medium' }
  };

  // Fixed layout for 5 meter cells (E7:I7)
  const meterCells = [
    { col: 'E7', value: lossesCalculationData.mainClient.meterNumber || '', bgColor: 'D9D9D9' }
  ];

  // Add subclient meter numbers (up to 5 columns total)
  const maxSubClients = 5;
  for (let i = 0; i < maxSubClients; i++) {
    const colChar = String.fromCharCode(69 + i + 1); // Start from F (70)
    const cellRef = `${colChar}7`;
    const value = i < lossesCalculationData.subClient.length ?
      lossesCalculationData.subClient[i].meterNumber || '' :
      '';

    meterCells.push({
      col: cellRef,
      value: value,
      bgColor: 'D9D9D9'
    });
  }

  // Apply formatting to all meter cells
  meterCells.forEach(cell => {
    const cellRef = summarySheet.getCell(cell.col);
    cellRef.value = cell.value;
    cellRef.font = { bold: true, size: 10, name: 'Times New Roman' };
    cellRef.alignment = { horizontal: 'center', vertical: 'middle' };
    cellRef.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: cell.bgColor }
    };
    cellRef.border = {
      top: { style: 'thin' },
      left: { style: 'medium' },
      bottom: { style: 'thin' },
      right: { style: 'medium' }
    };
  });

  // Row 8: Gross Drawl Units row
  const grossDrawlRow = summarySheet.getRow(8);
  grossDrawlRow.height = 45;

  // Gross Drawl Units label - merge A8:B8
  summarySheet.mergeCells('A8:B8');
  const grossDrawlLabelCell = summarySheet.getCell('A8');
  grossDrawlLabelCell.value = `Gross Drawl Units from ${lossesCalculationData.mainClient.mainClientDetail.subTitle}`;
  grossDrawlLabelCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  grossDrawlLabelCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
  grossDrawlLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '92D050' } // Green background
  };
  grossDrawlLabelCell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  // Gross Drawl Units value - C8
  const grossDrawlValueCell = summarySheet.getCell('C8');
  grossDrawlValueCell.value = displayExactValue(grossDrawlValue);
  grossDrawlValueCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  grossDrawlValueCell.alignment = { horizontal: 'center', vertical: 'middle' };
  grossDrawlValueCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '92D050' } // Green background
  };
  grossDrawlValueCell.border = {
    top: { style: 'thin' },
    left: { style: 'medium' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  // Voltage Level label - D8
  const voltageLabelCell = summarySheet.getCell('D8');
  voltageLabelCell.value = 'Voltage Level';
  voltageLabelCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  voltageLabelCell.alignment = { horizontal: 'right', vertical: 'middle', wrapText: true };
  voltageLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'D9D9D9' }
  };
  voltageLabelCell.border = {
    top: { style: 'thin' },
    left: { style: 'medium' },
    bottom: { style: 'thin' },
    right: { style: 'medium' }
  };

  // Fixed layout for voltage level cells (E8:I8)
  const voltageCells = [
    { col: 'E8', value: lossesCalculationData.mainClient.mainClientDetail.voltageLevel || '', bgColor: 'D9D9D9' }
  ];

  // Add subclient voltage levels (up to 5 columns total)
  for (let i = 0; i < maxSubClients; i++) {
    const colChar = String.fromCharCode(69 + i + 1); // Start from F (70)
    const cellRef = `${colChar}8`;
    const value = i < lossesCalculationData.subClient.length ?
      lossesCalculationData.subClient[i].voltageLevel || '' :
      '';

    voltageCells.push({
      col: cellRef,
      value: value,
      bgColor: 'D9D9D9'
    });
  }

  // Apply formatting to all voltage cells
  voltageCells.forEach(cell => {
    const cellRef = summarySheet.getCell(cell.col);
    cellRef.value = cell.value;
    cellRef.font = { size: 10, name: 'Times New Roman' };
    cellRef.alignment = { horizontal: 'center', vertical: 'middle' };
    cellRef.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: cell.bgColor }
    };
    cellRef.border = {
      top: { style: 'thin' },
      left: { style: 'medium' },
      bottom: { style: 'thin' },
      right: { style: 'medium' }
    };
  });

  // Row 9: Net Injected Units row
  const netInjectedRow = summarySheet.getRow(9);
  netInjectedRow.height = 45;

  // Net Injected Units label - merge A9:B9
  summarySheet.mergeCells('A9:B9');
  const netInjectedLabelCell = summarySheet.getCell('A9');
  netInjectedLabelCell.value = `Net Injected Units to ${lossesCalculationData.mainClient.mainClientDetail.subTitle}`;
  netInjectedLabelCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  netInjectedLabelCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
  netInjectedLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '92D050' } // Green background
  };
  netInjectedLabelCell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  // Net Injected Units value - C9
  const netInjectedValueCell = summarySheet.getCell('C9');
  netInjectedValueCell.value = displayExactValue(netInjectedValue);
  netInjectedValueCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  netInjectedValueCell.alignment = { horizontal: 'center', vertical: 'middle' };
  netInjectedValueCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '92D050' } // Green background
  };
  netInjectedValueCell.border = {
    top: { style: 'thin' },
    left: { style: 'medium' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  // CTPT Sr.No. label - D9
  const ctptLabelCell = summarySheet.getCell('D9');
  ctptLabelCell.value = 'CTPT Sr.No.';
  ctptLabelCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  ctptLabelCell.alignment = { horizontal: 'right', vertical: 'middle' };
  ctptLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'D9D9D9' }
  };
  ctptLabelCell.border = {
    top: { style: 'thin' },
    left: { style: 'medium' },
    bottom: { style: 'thin' },
    right: { style: 'medium' }
  };

  // Fixed layout for CTPT cells (E9:I9)
  const ctptCells = [
    { col: 'E9', value: lossesCalculationData.mainClient.mainClientDetail.ctptSrNo || '', bgColor: 'D9D9D9' }
  ];

  // Add subclient CTPT numbers (up to 5 columns total)
  for (let i = 0; i < maxSubClients; i++) {
    const colChar = String.fromCharCode(69 + i + 1); // Start from F (70)
    const cellRef = `${colChar}9`;
    const value = i < lossesCalculationData.subClient.length ?
      lossesCalculationData.subClient[i].ctptSrNo || '' :
      '';

    ctptCells.push({
      col: cellRef,
      value: value,
      bgColor: 'D9D9D9'
    });
  }

  // Apply formatting to all CTPT cells
  ctptCells.forEach(cell => {
    const cellRef = summarySheet.getCell(cell.col);
    cellRef.value = cell.value;
    cellRef.font = { size: 10, name: 'Times New Roman' };
    cellRef.alignment = { horizontal: 'center', vertical: 'middle' };
    cellRef.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: cell.bgColor }
    };
    cellRef.border = {
      top: { style: 'thin' },
      left: { style: 'medium' },
      bottom: { style: 'thin' },
      right: { style: 'medium' }
    };
  });

  // Row 10: Overall Percentage Distributions row
  const percentageRow = summarySheet.getRow(10);
  percentageRow.height = 45;

  // Overall Percentage Distributions label - merge A10:C10
  summarySheet.mergeCells('A10:C10');
  const percentageLabelCell = summarySheet.getCell('A10');
  percentageLabelCell.value = 'Overall Percentage Distributions';
  percentageLabelCell.font = { bold: true, size: 11, name: 'Times New Roman' };
  percentageLabelCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  percentageLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF00' } // Yellow background
  };
  percentageLabelCell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  // CT Ratio label - D10
  const ctRatioLabelCell = summarySheet.getCell('D10');
  ctRatioLabelCell.value = 'CT Ratio (A/A)';
  ctRatioLabelCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  ctRatioLabelCell.alignment = { horizontal: 'right', vertical: 'middle' };
  ctRatioLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'D9D9D9' }
  };
  ctRatioLabelCell.border = {
    top: { style: 'thin' },
    left: { style: 'medium' },
    bottom: { style: 'thin' },
    right: { style: 'medium' }
  };

  // Fixed layout for CT Ratio cells (E10:I10)
  const ctRatioCells = [
    { col: 'E10', value: lossesCalculationData.mainClient.mainClientDetail.ctRatio || '', bgColor: 'D9D9D9' }
  ];

  // Add subclient CT Ratios (up to 5 columns total)
  for (let i = 0; i < maxSubClients; i++) {
    const colChar = String.fromCharCode(69 + i + 1); // Start from F (70)
    const cellRef = `${colChar}10`;
    const value = i < lossesCalculationData.subClient.length ?
      lossesCalculationData.subClient[i].ctRatio || '' :
      '';

    ctRatioCells.push({
      col: cellRef,
      value: value,
      bgColor: 'D9D9D9'
    });
  }

  // Apply formatting to all CT Ratio cells
  ctRatioCells.forEach(cell => {
    const cellRef = summarySheet.getCell(cell.col);
    cellRef.value = cell.value;
    cellRef.font = { size: 10, name: 'Times New Roman' };
    cellRef.alignment = { horizontal: 'center', vertical: 'middle' };
    cellRef.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: cell.bgColor }
    };
    cellRef.border = {
      top: { style: 'thin' },
      left: { style: 'medium' },
      bottom: { style: 'thin' },
      right: { style: 'medium' }
    };
  });

  // Row 11: Discom row
  const discomRow = summarySheet.getRow(11);
  discomRow.height = 45;

  // Discom label - merge A11:B11
  summarySheet.mergeCells('A11:B11');
  const discomLabelCell = summarySheet.getCell('A11');
  discomLabelCell.value = 'DISCOM (%)';
  discomLabelCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  discomLabelCell.alignment = { horizontal: 'left', vertical: 'middle' };
  discomLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF00' } // Yellow background
  };
  discomLabelCell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  // Discom value - C11
  const discomValueCell = summarySheet.getCell('C11');
  discomValueCell.value = 'DGVCL 100 %';
  discomValueCell.font = { bold: true, size: 10, name: 'Times New Roman', color: { argb: 'FF0000' } };
  discomValueCell.alignment = { horizontal: 'center', vertical: 'middle' };
  discomValueCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF00' } // Yellow background
  };
  discomValueCell.border = {
    top: { style: 'thin' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' }
  };

  // PT Ratio label - D11
  const ptRatioLabelCell = summarySheet.getCell('D11');
  ptRatioLabelCell.value = 'PT Ratio (V/V)';
  ptRatioLabelCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  ptRatioLabelCell.alignment = { horizontal: 'right', vertical: 'middle' };
  ptRatioLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'D9D9D9' }
  };
  ptRatioLabelCell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  // Fixed layout for PT Ratio cells (E11:I11)
  const ptRatioCells = [
    { col: 'E11', value: lossesCalculationData.mainClient.mainClientDetail.ptRatio || '', bgColor: 'D9D9D9' }
  ];

  // Add subclient PT Ratios (up to 5 columns total)
  for (let i = 0; i < maxSubClients; i++) {
    const colChar = String.fromCharCode(69 + i + 1); // Start from F (70)
    const cellRef = `${colChar}11`;
    const value = i < lossesCalculationData.subClient.length ?
      lossesCalculationData.subClient[i].ptRatio || '' :
      '';

    ptRatioCells.push({
      col: cellRef,
      value: value,
      bgColor: 'D9D9D9'
    });
  }

  // Apply formatting to all PT Ratio cells
  ptRatioCells.forEach(cell => {
    const cellRef = summarySheet.getCell(cell.col);
    cellRef.value = cell.value;
    cellRef.font = { size: 10, name: 'Times New Roman' };
    cellRef.alignment = { horizontal: 'center', vertical: 'middle' };
    cellRef.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: cell.bgColor }
    };
    cellRef.border = {
      top: { style: 'thin' },
      left: { style: 'medium' },
      bottom: { style: 'thin' },
      right: { style: 'medium' }
    };
  });

  // Row 12: Units row
  const unitsRow = summarySheet.getRow(12);
  unitsRow.height = 45;

  // Units label - merge A12:B12
  summarySheet.mergeCells('A12:B12');
  const unitsLabelCell = summarySheet.getCell('A12');
  unitsLabelCell.value = 'Units (MWh)';
  unitsLabelCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  unitsLabelCell.alignment = { horizontal: 'left', vertical: 'middle' };
  unitsLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF00' } // Yellow background
  };
  unitsLabelCell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'medium' },
    right: { style: 'thin' }
  };

  // Units value - C12
  const unitsValueCell = summarySheet.getCell('C12');
  unitsValueCell.value = displayExactValue(netInjectedValue);
  unitsValueCell.font = { bold: true, size: 10, name: 'Times New Roman', color: { argb: 'FF0000' } };
  unitsValueCell.alignment = { horizontal: 'center', vertical: 'middle' };
  unitsValueCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF00' } // Yellow background
  };
  unitsValueCell.border = {
    top: { style: 'thin' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' }
  };

  // MF label - D12
  const mfLabelCell = summarySheet.getCell('D12');
  mfLabelCell.value = 'MF';
  mfLabelCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  mfLabelCell.alignment = { horizontal: 'right', vertical: 'middle' };
  mfLabelCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'D9D9D9' }
  };
  mfLabelCell.border = {
    top: { style: 'thin' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' }
  };

  // Fixed layout for MF cells (E12:I12)
  const mfCells = [
    {
      col: 'E12',
      value: lossesCalculationData.mainClient.mainClientDetail.mf || '5000', // Default 5000 for main client
      bgColor: 'D9D9D9'
    }
  ];

  // Add subclient MFs (up to 5 columns total)
  for (let i = 0; i < maxSubClients; i++) {
    const colChar = String.fromCharCode(69 + i + 1); // Start from F (70)
    const cellRef = `${colChar}12`;
    const value = i < lossesCalculationData.subClient.length ?
      lossesCalculationData.subClient[i].mf || '1000' : // Default 1000 for sub clients
      '';

    mfCells.push({
      col: cellRef,
      value: value,
      bgColor: 'D9D9D9'
    });
  }

  // Apply formatting to all MF cells
  mfCells.forEach(cell => {
    const cellRef = summarySheet.getCell(cell.col);
    cellRef.value = cell.value;
    cellRef.font = { size: 10, name: 'Times New Roman' };
    cellRef.alignment = { horizontal: 'center', vertical: 'middle' };
    cellRef.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: cell.bgColor }
    };
    cellRef.border = {
      top: { style: 'thin' },
      left: { style: 'medium' },
      bottom: { style: 'medium' },
      right: { style: 'medium' }
    };
  });
  // Add blank row
  summarySheet.getRow(13).height = 15;

  // Row 14: Table headers
  const tableHeaderRow = summarySheet.getRow(14);
  tableHeaderRow.height = 54;

  // Table headers
  const tableHeaders = [
    { cell: 'A14', value: 'Sr. No.', bgColor: 'D9D9D9', borderRight: 'thin', borderLeft: 'thin' }, // Special property for this cell
    { cell: 'B14', value: 'HT Consumer Name', bgColor: 'D9D9D9' },
    { cell: 'C14', value: 'HT Consumer No.', bgColor: 'D9D9D9' },
    { cell: 'D14', value: 'Wheeling Division Office/Location', bgColor: 'D9D9D9' },
    { cell: 'E14', value: 'Wheeling DISCOM', bgColor: 'D9D9D9' },
    {
      cell: 'F14',
      value: 'Project Capacity (kW) (AC)',
      bgColor: 'D9D9D9',
      borderRight: 'medium'  // Special property for this cell
    },
    { cell: 'G14', value: 'Share in Gross Injected Units to S/S (MWh)', bgColor: 'D9D9D9' },
    { cell: 'H14', value: 'Share in Gross Drawl Units from S/S (MWh)', bgColor: 'D9D9D9' },
    { cell: 'I14', value: 'Net Injected Units to S/S (MWh)', bgColor: 'D9D9D9' },
    { cell: 'J14', value: '% Weightage According to Gross Injecting', bgColor: 'D9D9D9' }
  ];

  // Apply table headers
  tableHeaders.forEach(header => {
    const cell = summarySheet.getCell(header.cell);
    cell.value = header.value;
    cell.font = { bold: true, size: 10, name: 'Times New Roman' };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: header.bgColor }
    };
    cell.border = {
      top: { style: 'thin' },
      left: { style: header.borderLeft || 'thin' },  // Use specified or default to thin
      bottom: { style: 'thin' },
      right: { style: header.borderRight || 'thin' }  // Use specified or default to thin
    };
  });

  // Add medium borders around the entire header row (outer borders)
  const headerRow = 14;
  const firstHeaderCol = 'A';
  const lastHeaderCol = 'J';

  // Left border for first column
  const leftCell = summarySheet.getCell(`${firstHeaderCol}${headerRow}`);
  leftCell.border = {
    ...leftCell.border,
    left: { style: 'medium' }
  };

  // Right border for last column
  const rightCell = summarySheet.getCell(`${lastHeaderCol}${headerRow}`);
  rightCell.border = {
    ...rightCell.border,
    right: { style: 'medium' }
  };

  // Top and bottom borders for all header cells
  for (let col = 1; col <= 10; col++) {
    const colChar = String.fromCharCode(64 + col);
    const cell = summarySheet.getCell(`${colChar}${headerRow}`);
    cell.border = {
      ...cell.border,
      top: { style: 'medium' },
      bottom: { style: 'medium' }
    };
  }
  // Add data rows for each subclient
  let rowNum = 15;
  let globalIndex = 1; // Initialize a global counter for sequential numbering

  lossesCalculationData.subClient.forEach((subClient) => {
    const subClientData = subClient.subClientsData || {};

    // Check if this subclient has partclients
    if (subClientData.partclient && subClientData.partclient.length > 0) {
      // Add each partclient as a separate row
      subClientData.partclient.forEach((partClient, partIndex) => {
        const currentRowNum = rowNum++;
        const row = summarySheet.getRow(currentRowNum);
        row.height = 60;

        const grossInjection = partClient.grossInjectionMWHAfterLosses || 0;
        const drawl = partClient.drawlMWHAfterLosses || 0;
        const netInjection = grossInjection + drawl;
        const weightage = grossInjectedValue > 0 ? (grossInjection / grossInjectedValue) * 100 : 0;

        // Sr. No. - Use globalIndex and increment it
        summarySheet.getCell(`A${currentRowNum}`).value = globalIndex++;
        summarySheet.getCell(`A${currentRowNum}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        summarySheet.getCell(`A${currentRowNum}`).border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };

        // HT Consumer Name with sharing percentage - LEFT ALIGNED
        // HT Consumer Name with sharing percentage - LEFT ALIGNED
        summarySheet.getCell(`B${currentRowNum}`).value = {
          richText: [
            { text: `${subClient.name.toUpperCase()} - Unit-${partIndex + 1}`, font: { size: 10, name: 'Times New Roman' } },
            { text: ` (${partClient.sharingPercentage}% of Sharing OA)`, font: { size: 10, name: 'Times New Roman' } }
          ]
        };
        summarySheet.getCell(`B${currentRowNum}`).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
        summarySheet.getCell(`B${currentRowNum}`).border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
        summarySheet.getCell(`B${currentRowNum}`).font = { size: 10, name: 'Times New Roman' };

        // HT Consumer No.
        summarySheet.getCell(`C${currentRowNum}`).value = partClient.consumerNo || '';
        summarySheet.getCell(`C${currentRowNum}`).alignment = { horizontal: 'center', vertical: 'middle' };

        // Wheeling Division Office/Location
        summarySheet.getCell(`D${currentRowNum}`).value = subClient.divisionName || '';
        summarySheet.getCell(`D${currentRowNum}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

        // Wheeling Discom
        summarySheet.getCell(`E${currentRowNum}`).value = subClient.discom || '';
        summarySheet.getCell(`E${currentRowNum}`).alignment = { horizontal: 'center', vertical: 'middle' };

        // Project Capacity (kW) (AC) - only show for first partclient
        if (partIndex === 0) {
          summarySheet.getCell(`F${currentRowNum}`).value = subClient.acCapacityKw || '';
          summarySheet.getCell(`F${currentRowNum}`).alignment = { horizontal: 'center', vertical: 'middle' };
          // Add medium right border to Project Capacity cell
          summarySheet.getCell(`F${currentRowNum}`).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'medium' }  // Medium right border
          };

          // Merge capacity cells vertically for all partclients
          if (subClientData.partclient.length > 1) {
            summarySheet.mergeCells(`F${currentRowNum}:F${currentRowNum + subClientData.partclient.length - 1}`);
            // Also apply the medium right border to merged cells
            for (let i = 0; i < subClientData.partclient.length; i++) {
              summarySheet.getCell(`F${currentRowNum + i}`).border = {
                top: i === 0 ? { style: 'thin' } : { style: 'none' },
                left: { style: 'thin' },
                bottom: i === subClientData.partclient.length - 1 ? { style: 'thin' } : { style: 'none' },
                right: { style: 'medium' }  // Medium right border for all merged cells
              };
            }
          }
        }

        // Apply base formatting to all other cells (except column B and F which are already set)
        for (let col = 3; col <= 10; col++) { // Start from column C (3)
          if (col !== 6) { // Skip column F (6) as it's already handled
            const cell = row.getCell(col);
            cell.font = { size: 10, name: 'Times New Roman' };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };
          }
        }

        // Now apply bold formatting to specific cells
        const boldCells = ['G', 'H', 'I', 'J'];
        boldCells.forEach(col => {
          const cell = summarySheet.getCell(`${col}${currentRowNum}`);
          cell.font = { ...cell.font, bold: true };
        });

        // Set values for bold cells
        summarySheet.getCell(`G${currentRowNum}`).value = displayExactValue(grossInjection);
        summarySheet.getCell(`H${currentRowNum}`).value = displayExactValue(drawl);
        summarySheet.getCell(`I${currentRowNum}`).value = displayExactValue(netInjection);
        // summarySheet.getCell(`J${currentRowNum}`).value = `${displayExactValue(weightage, 2)} %`;
        summarySheet.getCell(`J${currentRowNum}`).value = `${Number(weightage).toFixed(2)} %`;

      });
    } else {
      // Regular subclient without partclients
      const currentRowNum = rowNum++;
      const row = summarySheet.getRow(currentRowNum);
      row.height = 60;

      const grossInjection = subClientData.grossInjectionMWHAfterLosses || 0;
      const drawl = subClientData.drawlMWHAfterLosses || 0;
      const netInjection = subClientData.netInjectionMWHAfterLosses || 0;
      const weightage = (netInjection / lossesCalculationData.mainClient.netInjectionMWH) * 100 || 0; // Use main client's net injection for weightage calculation
      // const weightage = grossInjectedValue > 0 ? (grossInjection / grossInjectedValue) * 100 : 0;

      // Sr. No. - Use globalIndex and increment it
      summarySheet.getCell(`A${currentRowNum}`).value = globalIndex++;
      summarySheet.getCell(`A${currentRowNum}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      summarySheet.getCell(`A${currentRowNum}`).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };

      // HT Consumer Name - LEFT ALIGNED
      summarySheet.getCell(`B${currentRowNum}`).value = `${subClient.name.toUpperCase()}`;
      summarySheet.getCell(`B${currentRowNum}`).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      summarySheet.getCell(`B${currentRowNum}`).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
      summarySheet.getCell(`B${currentRowNum}`).font = { size: 10, name: 'Times New Roman' };

      // HT Consumer No.
      summarySheet.getCell(`C${currentRowNum}`).value = subClient.consumerNo || '';
      summarySheet.getCell(`C${currentRowNum}`).alignment = { horizontal: 'center', vertical: 'middle' };

      // Wheeling Division Office/Location
      summarySheet.getCell(`D${currentRowNum}`).value = subClient.divisionName || '';
      summarySheet.getCell(`D${currentRowNum}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

      // Wheeling Discom
      summarySheet.getCell(`E${currentRowNum}`).value = subClient.discom;
      summarySheet.getCell(`E${currentRowNum}`).alignment = { horizontal: 'center', vertical: 'middle' };

      // Project Capacity (kW) (AC) with medium right border
      summarySheet.getCell(`F${currentRowNum}`).value = subClient.acCapacityKw || '';
      summarySheet.getCell(`F${currentRowNum}`).alignment = { horizontal: 'center', vertical: 'middle' };
      summarySheet.getCell(`F${currentRowNum}`).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'medium' }  // Medium right border
      };

      // Apply base formatting to all other cells (except column B and F which are already set)
      for (let col = 3; col <= 10; col++) { // Start from column C (3)
        if (col !== 6) { // Skip column F (6) as it's already handled
          const cell = row.getCell(col);
          cell.font = { size: 10, name: 'Times New Roman' };
          cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };
        }
      }

      // Now apply bold formatting to specific cells
      const boldCells = ['G', 'H', 'I', 'J'];
      boldCells.forEach(col => {
        const cell = summarySheet.getCell(`${col}${currentRowNum}`);
        cell.font = { ...cell.font, bold: true };
      });

      // Set values for bold cells
      summarySheet.getCell(`G${currentRowNum}`).value = displayExactValue(grossInjection);
      summarySheet.getCell(`H${currentRowNum}`).value = displayExactValue(drawl);
      summarySheet.getCell(`I${currentRowNum}`).value = displayExactValue(netInjection);
      // summarySheet.getCell(`J${currentRowNum}`).value = `${displayExactValue(weightage, 2)} %`;
      summarySheet.getCell(`J${currentRowNum}`).value = `${Number(weightage).toFixed(2)} %`;
    }
  });

  // Add blank row where main client row would have been
  const blankRowNum = rowNum++;
  const blankRow = summarySheet.getRow(blankRowNum);
  blankRow.height = 15;

  // Apply minimal formatting to blank row
  for (let col = 1; col <= 10; col++) {
    const cell = blankRow.getCell(col);
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: col === 6 ? 'medium' : 'thin' } // Column F is the 6th column
    };
  }

  // Add total row (now comes after blank row)
  const totalRowNum = blankRowNum + 1;
  const totalRow = summarySheet.getRow(totalRowNum);
  totalRow.height = 28;

  // Merge and add "Total" label
  summarySheet.mergeCells(`A${totalRowNum}:E${totalRowNum}`);
  const totalLabelCells = summarySheet.getCell(`A${totalRowNum}`);
  totalLabelCells.value = 'Total';
  totalLabelCells.font = { bold: true, size: 12, name: 'Times New Roman' };
  totalLabelCells.alignment = { horizontal: 'center', vertical: 'middle' };
  totalLabelCells.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'D9D9D9' }
  };
  totalLabelCells.border = {
    top: { style: 'medium' }, // Changed to medium
    left: { style: 'thin' },
    bottom: { style: 'medium' }, // Changed to medium
    right: { style: 'thin' }
  };

  // Calculate totals from all subclients and partclients
  let totalCapacity = 0;
  let totalGrossInjected = 0;
  let totalGrossDrawl = 0;

  lossesCalculationData.subClient.forEach(subClient => {
    const subClientData = subClient.subClientsData || {};

    if (subClientData.partclient && subClientData.partclient.length > 0) {
      // Only count capacity once for the parent subclient
      totalCapacity += subClient.acCapacityKw || 0;

      // Sum all partclient values
      subClientData.partclient.forEach(partClient => {
        totalGrossInjected += partClient.grossInjectionMWHAfterLosses || 0;
        totalGrossDrawl += partClient.drawlMWHAfterLosses || 0;
      });
    } else {
      totalCapacity += subClient.acCapacityKw || 0;
      totalGrossInjected += subClientData.grossInjectionMWHAfterLosses || 0;
      totalGrossDrawl += subClientData.drawlMWHAfterLosses || 0;
    }
  });

  const totalNetInjected = totalGrossInjected + totalGrossDrawl;

  // Project Capacity total
  const fCell = summarySheet.getCell(`F${totalRowNum}`);
  fCell.value = totalCapacity;
  fCell.border = {  // Added specific border for F cell
    top: { style: 'medium' },
    left: { style: 'thin' },
    bottom: { style: 'medium' },
    right: { style: 'medium' } // Medium right border
  };

  // Gross Injected total
  summarySheet.getCell(`G${totalRowNum}`).value = displayExactValue(totalGrossInjected);

  // Gross Drawl total
  summarySheet.getCell(`H${totalRowNum}`).value = displayExactValue(totalGrossDrawl);

  // Net Injected total
  summarySheet.getCell(`I${totalRowNum}`).value = displayExactValue(totalNetInjected);

  // % Weightage total
  summarySheet.getCell(`J${totalRowNum}`).value = '100%';

  // Apply formatting to total cells
  for (let col = 6; col <= 10; col++) {
    const cell = totalRow.getCell(col);
    cell.font = { bold: true, size: 12, name: 'Times New Roman' };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'D9D9D9' }
    };
    cell.border = {
      top: { style: 'medium' }, // Changed to medium
      left: { style: 'thin' },
      bottom: { style: 'medium' }, // Changed to medium
      right: { style: col === 6 ? 'medium' : 'thin' } // Medium right border only for F (col 6)
    };
  }

  // Add blank row
  summarySheet.getRow(totalRowNum + 1).height = 15;

  // Add note rows
  const noteRow1 = totalRowNum + 2; // Row 22 in your example
  const noteRow2 = totalRowNum + 3; // Row 23 in your example

  // Row 22 - First note line
  summarySheet.getCell(`A${noteRow1}`).value = "Note:";
  summarySheet.getCell(`A${noteRow1}`).font = { bold: true, italic: true, size: 10, name: 'Times New Roman' };
  summarySheet.mergeCells(`B${noteRow1}:J${noteRow1}`);
  const noteCell1 = summarySheet.getCell(`B${noteRow1}`);
  noteCell1.value = {
    richText: [
      { text: "1) All Units are in ", font: { bold: true, italic: true, size: 10, name: 'Times New Roman' } },
      { text: "MWH", font: { bold: true, italic: true, size: 10, name: 'Times New Roman', color: { argb: 'FF0000' } } }
    ]
  };

  // Row 23 - Second note line
  summarySheet.getCell(`A${noteRow2}`).value = ""; // Empty cell in column A
  summarySheet.mergeCells(`B${noteRow2}:J${noteRow2}`);
  const noteCell2 = summarySheet.getCell(`B${noteRow2}`);
  noteCell2.value = {
    richText: [
      { text: "2) ", font: { bold: true, italic: true, size: 10, name: 'Times New Roman' } },
      { text: `${lossesCalculationData.mainClient.meterNumber}`, font: { italic: true, size: 10, name: 'Times New Roman', bold: true } },
      { text: ` is the Grossing Meter at ${lossesCalculationData.mainClient.mainClientDetail.subTitle} S/S End`, font: { bold: true, italic: true, size: 10, name: 'Times New Roman' } }
    ]
  };

  // Apply common formatting to both rows
  [noteRow1, noteRow2].forEach(row => {
    for (let col = 1; col <= 10; col++) {
      const cell = summarySheet.getCell(`${String.fromCharCode(64 + col)}${row}`);
      cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      cell.border = {
        top: { style: 'medium' },
        left: { style: 'medium' },
        bottom: { style: 'medium' },
        right: { style: 'medium' }
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
      left: { style: 'medium' }
    };

    // Right border (column J)
    const rightCell = summarySheet.getCell(`J${row}`);
    rightCell.border = {
      ...rightCell.border,
      right: { style: 'medium' }
    };
  }

  // Apply top border to first row
  for (let col = 1; col <= 10; col++) {
    const colChar = String.fromCharCode(64 + col);
    const cell = summarySheet.getCell(`${colChar}${firstDataRow}`);
    cell.border = {
      ...cell.border,
      top: { style: 'medium' }
    };
  }

  // Apply bottom border to last row
  for (let col = 1; col <= 10; col++) {
    const colChar = String.fromCharCode(64 + col);
    const cell = summarySheet.getCell(`${colChar}${lastDataRow}`);
    cell.border = {
      ...cell.border,
      bottom: { style: 'medium' }
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
    `B${noteRow1}:J${noteRow1}`, `B${noteRow2}:J${noteRow2}` // Note rows
  ];

  mergedAreas.forEach(merge => {
    const cell = summarySheet.getCell(merge.split(':')[0]);
    cell.border = {
      ...cell.border,
      top: { style: 'medium' },
      left: { style: 'medium' },
      bottom: { style: 'medium' },
      right: { style: 'medium' }
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
    top: { style: 'medium', color: { argb: '000000' } },
    left: { style: 'medium', color: { argb: '000000' } },
    bottom: { style: 'medium', color: { argb: '000000' } },
    right: { style: 'medium', color: { argb: '000000' } }
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
    top: { style: 'medium', color: { argb: '000000' } },
    left: { style: 'medium', color: { argb: '000000' } },
    bottom: { style: 'medium', color: { argb: '000000' } },
    right: { style: 'medium', color: { argb: '000000' } }
  };

  // Date range from H3:J3
  masterdataSheet.mergeCells('H3:J3');
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
    top: { style: 'medium', color: { argb: '000000' } },
    left: { style: 'medium', color: { argb: '000000' } },
    bottom: { style: 'medium', color: { argb: '000000' } },
    right: { style: 'medium', color: { argb: '000000' } }
  };

  masterdataSheet.getColumn('H').width = 14;
  masterdataSheet.getColumn('I').width = 14;
  masterdataSheet.getColumn('J').width = 14;

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
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' }
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
          left: { style: 'medium' },
          bottom: { style: 'thin' },
          right: { style: 'medium' }
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
          left: { style: 'medium' },
          bottom: { style: 'thin' },
          right: { style: 'medium' }
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
          left: { style: 'medium' },
          bottom: { style: 'thin' },
          right: { style: 'medium' }
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
          left: { style: 'medium' },
          bottom: { style: 'thin' },
          right: { style: 'medium' }
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
          top: { style: 'medium' },
          left: { style: 'medium' },
          bottom: { style: 'medium' },
          right: { style: 'medium' }
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
        left: { style: 'medium' },
        bottom: { style: 'thin' },
        right: { style: 'medium' }
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
        left: { style: 'medium' },
        bottom: { style: 'thin' },
        right: { style: 'medium' }
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
        left: { style: 'medium' },
        bottom: { style: 'thin' },
        right: { style: 'medium' }
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
        left: { style: 'medium' },
        bottom: { style: 'thin' },
        right: { style: 'medium' }
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
        left: { style: 'medium' },
        bottom: { style: 'thin' },
        right: { style: 'medium' }
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
        top: { style: 'medium' },
        left: { style: 'medium' },
        bottom: { style: 'medium' },
        right: { style: 'medium' }
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
  totalCell.value = 'Total Share (MWh)';
  totalCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  totalCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  totalCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFFF' }
  };
  totalCell.border = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' }
  };
  masterdataSheet.getColumn(totalCol).width = 22;

  // CHECK-SUM (merged from row 5 to 8)
  masterdataSheet.mergeCells(`${checkSumCol}5:${checkSumCol}8`);
  const checkCell = masterdataSheet.getCell(`${checkSumCol}5`);
  checkCell.value = 'CHECK-SUM';
  checkCell.font = { bold: true, size: 10, name: 'Times New Roman' };
  checkCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  checkCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFFF' }
  };
  checkCell.border = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' }
  };
  masterdataSheet.getColumn(checkSumCol).width = 22;

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

        // Add CHECK-SUM
        values.push("0.000");

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

        values.push("0.000");
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

        values.push("0.000");
        return values;
      },
      format: value => value.toFixed(3)
    }
  ];

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
    values.forEach((value, colIndex) => {
      const colLetter = String.fromCharCode(68 + colIndex); // D, E, F...
      const cell = masterdataSheet.getCell(`${colLetter}${dataRow}`);
      cell.value = typeof value === 'number' ? row.format(value) : value;
      cell.font = { size: 10, bold: true, name: 'Times New Roman' };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };

      // Apply color to value cells based on column
      if (colLetter === 'D') {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '92D050' } // Green for main client          
        };
      } else if (colIndex >= 1 && colIndex < values.length - 2) {
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

        // CHECK-SUM value (original styling)
        const checkSumCell = masterdataSheet.getCell(`${checkSumCol}${rowIndex}`);
        const checkSumValue = subClientsSum - mainValue;
        checkSumCell.value = checkSumValue;
        checkSumCell.numFmt = '0.0';
        checkSumCell.font = { size: 10, name: 'Times New Roman' };
        checkSumCell.alignment = { horizontal: 'center', vertical: 'middle' };
        checkSumCell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };

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
  // ===== ADD MEDIUM BORDER AROUND MAIN DATA SECTION (ROWS 5-12) =====
  const lastDataCol = checkSumCol; // This is the CHECK-SUM column

  // Apply medium border around the entire section
  for (let row = 5; row <= 12; row++) {
    for (let col = 1; col <= lastDataCol.charCodeAt(0) - 64; col++) {
      const colLetter = String.fromCharCode(64 + col);
      const cell = masterdataSheet.getCell(`${colLetter}${row}`);

      // Special handling for COMBINED (D), Total (totalCol), and CHECK-SUM (checkSumCol) columns
      const isSpecialColumn = colLetter === 'D' || colLetter === totalCol || colLetter === checkSumCol;

      // Determine border style based on position
      const borderStyles = {
        top: row === 5 || isSpecialColumn ? 'medium' : 'thin',
        bottom: row === 12 ? 'medium' : 'thin',
        left: col === 1 ? 'medium' : 'medium',
        right: colLetter === lastDataCol ? 'medium' : 'thin'
      };

      // For special columns, ensure top border is medium from rows 5-12
      if (isSpecialColumn && row >= 5 && row <= 12) {
        borderStyles.top = 'medium';
      }

      cell.border = {
        top: { style: borderStyles.top, color: { argb: '000000' } },
        left: { style: borderStyles.left, color: { argb: '000000' } },
        bottom: { style: borderStyles.bottom, color: { argb: '000000' } },
        right: { style: borderStyles.right, color: { argb: '000000' } }
      };
    }
  }

  // Special case for the blank row (row 9) - ensure it has medium borders on all sides
  for (let col = 1; col <= lastDataCol.charCodeAt(0) - 64; col++) {
    const colLetter = String.fromCharCode(64 + col);
    const cell = masterdataSheet.getCell(`${colLetter}9`);

    cell.border = {
      top: { style: 'medium', color: { argb: '000000' } },
      left: { style: 'medium', color: { argb: '000000' } },
      bottom: { style: 'medium', color: { argb: '000000' } },
      right: { style: 'medium', color: { argb: '000000' } }
    };
  }

  // Explicitly set medium top borders for special columns (rows 5-12)
  ['D', totalCol, checkSumCol].forEach(colLetter => {
    for (let row = 5; row <= 12; row++) {
      const cell = masterdataSheet.getCell(`${colLetter}${row}`);
      cell.border = {
        ...cell.border,
        top: { style: 'medium', color: { argb: '000000' } }
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
    top: { style: 'medium', color: { argb: '000000' } },
    left: { style: 'medium', color: { argb: '000000' } },
    bottom: { style: 'medium', color: { argb: '000000' } },
    right: { style: 'medium', color: { argb: '000000' } }
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
    top: { style: 'medium', color: { argb: '000000' } },
    left: { style: 'medium', color: { argb: '000000' } },
    bottom: { style: 'medium', color: { argb: '000000' } },
    right: { style: 'medium', color: { argb: '000000' } }
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
    top: { style: 'medium', color: { argb: '000000' } },
    left: { style: 'medium', color: { argb: '000000' } },
    bottom: { style: 'medium', color: { argb: '000000' } },
    right: { style: 'medium', color: { argb: '000000' } }
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
      top: { style: 'medium' },
      left: { style: 'medium' },
      bottom: { style: 'medium' },
      right: { style: 'medium' }
    };
  });

  // ===== SUB-HEADER ROW =====
  // Sub-header row
  const subHeaders = [
    { cell: 'A5', value: 'Sr. No.' },
    { cell: 'B5', value: 'HT Consumer Name' },
    { cell: 'C5', value: 'Gross Injected Units (MWH)' },
    { cell: 'D5', value: 'Overall Gross Injected Units(MWH)' },
    { cell: 'E5', value: 'Gross Drawl Units (MWH)' },
    { cell: 'F5', value: 'Overall Gross Drawl Units (MWH)' },
    { cell: 'G5', value: 'Gross Received Units at S/S (MWH)' },
    { cell: 'H5', value: 'Net Drawl Units from S/S (MWH)' },
    { cell: 'I5', value: 'Difference in Injected Units, Plant End to S/S End (MWH)' },
    { cell: 'J5', value: 'Difference in Drawl Units, S/S End to Plant End (MWH)' },
    { cell: 'K5', value: '% Weightage According to Gross Injecting Units' },
    { cell: 'L5', value: '% Weightage According to Gross Drawl Units' },
    { cell: 'M5', value: 'Losses in Injected Units (MWH)' },
    { cell: 'N5', value: 'in %' },
    { cell: 'O5', value: 'Losses in Drawl Units (MWH)' },
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
  worksheet.getRow(5).height = 90; // Set height for header row

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
        top: { style: 'medium' },
        left: { style: 'medium' },
        bottom: { style: 'medium' },
        right: { style: 'medium' }
      };
    });

    // Add data rows for each sub client
    dataRowsLosses.forEach((rowData, index) => {
      const rowIndex = index + 6; // Starting from row 6
      const row = worksheet.getRow(rowIndex);
      row.height = 80; // Set row height

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
    // ===== ADD MEDIUM BORDER AROUND ENTIRE TABLE =====
    const firstDataRow = 4;
    const lastDataRows = 5 + dataRowsLosses.length + 1; // Includes header and total row
    const lastCol = 'P';

    // Apply medium border around the entire table
    for (let row = firstDataRow; row <= lastDataRows; row++) {
      for (let col = 1; col <= lastCol.charCodeAt(0) - 64; col++) {
        const colLetter = String.fromCharCode(64 + col);
        const cell = worksheet.getCell(`${colLetter}${row}`);

        // Determine border style based on position
        const borderStyles = {
          top: row === firstDataRow ? 'medium' : 'thin',
          bottom: row === lastDataRows ? 'medium' : 'thin',
          left: col === 1 ? 'medium' : 'thin',
          right: colLetter === lastCol ? 'medium' : 'thin'
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
        top: { style: 'medium', color: { argb: '000000' } },
        left: { style: 'medium', color: { argb: '000000' } },
        bottom: { style: 'medium', color: { argb: '000000' } },
        right: { style: 'medium', color: { argb: '000000' } }
      };
    });

    // Special handling for total row to ensure all borders are medium
    for (let col = 1; col <= lastCol.charCodeAt(0) - 64; col++) {
      const colLetter = String.fromCharCode(64 + col);
      const cell = worksheet.getCell(`${colLetter}${lastDataRows}`);

      cell.border = {
        top: { style: 'medium', color: { argb: '000000' } },
        left: { style: col === 1 ? 'medium' : 'thin', color: { argb: '000000' } },
        bottom: { style: 'medium', color: { argb: '000000' } },
        right: { style: colLetter === lastCol ? 'medium' : 'thin', color: { argb: '000000' } }
      };
    }

    // ===== UPDATE SPECIFIC COLUMN BORDERS FOR SUBCLIENT DATA ROWS =====
    const firstSubclientDataRow = 6;
    const lastSubclientDataRow = 5 + dataRowsLosses.length; // Last row of subclient data

    // Columns that need medium left border (Gross Injected Units - C)
    const columnsWithMediumLeftBorder = ['C'];

    // Columns that need medium right border (Overall Gross Drawl - F, Net Drawl - H, 
    // Difference in Drawl - J, % Weightage Drawl - L)
    const columnsWithMediumRightBorder = ['F', 'H', 'J', 'L'];

    // Apply medium borders to specified columns for all data rows
    for (let currentRow = firstSubclientDataRow; currentRow <= lastSubclientDataRow; currentRow++) {
      // Medium left borders
      columnsWithMediumLeftBorder.forEach(columnLetter => {
        const targetCell = worksheet.getCell(`${columnLetter}${currentRow}`);
        targetCell.border = {
          ...targetCell.border,
          left: { style: 'medium', color: { argb: '000000' } }
        };
      });

      // Medium right borders
      columnsWithMediumRightBorder.forEach(columnLetter => {
        const targetCell = worksheet.getCell(`${columnLetter}${currentRow}`);
        targetCell.border = {
          ...targetCell.border,
          right: { style: 'medium', color: { argb: '000000' } }
        };
      });
    }

    // Also update the header row (row 5) for these columns
    columnsWithMediumLeftBorder.forEach(columnLetter => {
      const headerCell = worksheet.getCell(`${columnLetter}5`);
      headerCell.border = {
        ...headerCell.border,
        left: { style: 'medium', color: { argb: '000000' } }
      };
    });

    columnsWithMediumRightBorder.forEach(columnLetter => {
      const headerCell = worksheet.getCell(`${columnLetter}5`);
      headerCell.border = {
        ...headerCell.border,
        right: { style: 'medium', color: { argb: '000000' } }
      };
    });

    // Update the total row borders for these columns
    const totalRowNumber = 5 + dataRowsLosses.length + 1;
    columnsWithMediumLeftBorder.forEach(columnLetter => {
      const totalRowCell = worksheet.getCell(`${columnLetter}${totalRowNumber}`);
      totalRowCell.border = {
        ...totalRowCell.border,
        left: { style: 'medium', color: { argb: '000000' } }
      };
    });

    columnsWithMediumRightBorder.forEach(columnLetter => {
      const totalRowCell = worksheet.getCell(`${columnLetter}${totalRowNumber}`);
      totalRowCell.border = {
        ...totalRowCell.border,
        right: { style: 'medium', color: { argb: '000000' } }
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
      top: { style: "medium", color: { argb: '000000' } },
      left: { style: "medium", color: { argb: '000000' } },
      bottom: { style: "medium", color: { argb: '000000' } },
      right: { style: "medium", color: { argb: '000000' } }
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
      top: { style: "medium", color: { argb: '000000' } },
      left: { style: "medium", color: { argb: '000000' } },
      bottom: { style: "medium", color: { argb: '000000' } },
      right: { style: "medium", color: { argb: '000000' } }
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
        top: { style: "medium", color: { argb: '000000' } },
        left: { style: "medium", color: { argb: '000000' } },
        bottom: { style: "medium", color: { argb: '000000' } },
        right: { style: "medium", color: { argb: '000000' } }
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
      top: { style: "medium" },
      left: { style: "medium" },
      bottom: { style: "medium" },
      right: { style: "medium" },
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
      top: { style: "medium" },
      left: { style: "medium" },
      bottom: { style: "medium" },
      right: { style: "medium" },
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
          top: { style: "medium" },
          left: { style: "medium" },
          bottom: { style: "medium" },
          right: { style: "medium" },
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
          top: { style: "medium" },
          left: { style: "medium" },
          bottom: { style: "medium" },
          right: { style: "medium" },
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
