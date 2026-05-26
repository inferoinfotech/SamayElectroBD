const fs = require('fs');
const path = require('path');

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

const formatDateDDMMYYYY = (date) => {
  const d = String(date.getDate()).padStart(2, '0');
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const y = date.getFullYear();
  return `${d}/${m}/${y}`;
};

/** Week 1 = Jan 1–7, week 2 = Jan 8–14, etc. */
const getWeekDateRange = (year, week) => {
  const w = parseInt(week, 10);
  const y = parseInt(year, 10);
  const start = new Date(y, 0, 1 + (w - 1) * 7);
  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  return {
    start: formatDateDDMMYYYY(start),
    end: formatDateDDMMYYYY(end),
  };
};

const getDefaultSignatureHtml = () => {
  const backendUrl = process.env.BACKEND_URL || 'http://localhost:8000';
  const logoExtensions = ['png', 'jpg', 'jpeg', 'svg'];
  const publicPath = path.join(__dirname, '../../public/images');
  let logoFileName = 'logo.png';

  if (fs.existsSync(publicPath)) {
    for (const ext of logoExtensions) {
      if (fs.existsSync(path.join(publicPath, `logo.${ext}`))) {
        logoFileName = `logo.${ext}`;
        break;
      }
    }
  }

  const logoUrl = `${backendUrl}/public/images/${logoFileName}`;

  return `<div style="font-family: Arial, sans-serif; color: #333;">
<p>Regards,</p>
<p><strong>Manish Thummar</strong><br/>
Mobile : +91 82385 85535</p>
<div style="margin: 20px 0;">
    <img src="${logoUrl}" alt="Samay Electro Service" style="max-width: 300px;" />
</div>
<p><strong>A-203, 2nd Floor, Dev Residency,<br/>
Near Verachha Co-Op. Bank, Yogichowk, Punagam,<br/>
Surat-395010, Gujarat, India.</strong></p>
<p>E-mail: <a href="mailto:info@samayelectro.com">info@samayelectro.com</a> | <a href="mailto:admin@samayelectro.com">admin@samayelectro.com</a></p>
<p><strong>MSME No.: UDYAM-GJ-22-0293351</strong></p>
<p><strong>GSTIN : 24AJTPT1949D1ZU</strong></p>
<p><strong>Working Hours (IST): 10:00 am to 6:00 pm, Sunday Week off</strong></p>
<p style="color: #4CAF50; font-size: 12px;"><em>please consider the environment before printing this email</em></p>
</div>`;
};

const buildAbatBody = (introParagraph, signature) =>
  `<p><strong>To, SLDC, Baroda</strong></p>
<p>Dear Sir/Madam,</p>
<p>${introParagraph}</p>
<p><strong>Project Details:</strong><br/>
Project Name: <strong>{{CLIENT_NAME}}</strong> – Solar<br/>
Location: <strong>{{FEEDER_NAME}}</strong> GETCO SS End</p>
<p><strong>Meter Details:</strong></p>
<p>{{MAIN_METER_NO}} – Main Meter<br/>
{{CHECK_METER_NO}} – Check Meter</p>
<p>Kindly acknowledge receipt of the submitted data.</p>
${signature}`;

const getDefaultTemplates = () => {
  const signature = getDefaultSignatureHtml();

  return {
    weekly: {
      subject:
        '{{CLIENT_NAME}} – Submission of Weekly ABT Meter Data – Week {{WEEK_NO}} ({{WEEK_START_DATE}} to {{WEEK_END_DATE}})',
      body: buildAbatBody(
        'Please find attached the ABT meter data for the <strong>{{CLIENT_NAME}}</strong> Solar Power Project, recorded at the <strong>{{FEEDER_NAME}}</strong> GETCO Substation end for <strong>Week {{WEEK_NO}} ({{WEEK_START_DATE}} to {{WEEK_END_DATE}})</strong>.',
        signature
      ),
    },
    monthly: {
      subject:
        '{{CLIENT_NAME}} – Submission of Monthly ABT Meter Data – {{MONTH_NAME}} {{YEAR}}',
      body: buildAbatBody(
        'Please find attached the ABT meter data for the <strong>{{CLIENT_NAME}}</strong> Solar Power Project, recorded at the <strong>{{FEEDER_NAME}}</strong> GETCO Substation end for <strong>{{MONTH_NAME}} – {{YEAR}}</strong>.',
        signature
      ),
    },
    general: {
      subject: '{{CLIENT_NAME}} – Monthly Documents – {{MONTH_NAME}} {{YEAR}}',
      body: `<p>PFA</p>
<p>Dear Sir/Madam,</p>
<p>Please find attached the documents for <strong>{{CLIENT_NAME}}</strong> for <strong>{{MONTH_NAME}} {{YEAR}}</strong>.</p>
<p>Kindly acknowledge receipt.</p>
${signature}`,
    },
  };
};

const applyEmailTemplate = (text, variables = {}) => {
  if (!text) return '';
  return String(text).replace(/\{\{(\w+)\}\}/g, (_, key) => {
    const value = variables[key];
    return value !== undefined && value !== null ? String(value) : `{{${key}}}`;
  });
};

const buildEmailVariables = ({ client, recipient, period, sendType }) => {
  const vars = {
    CLIENT_NAME:
      recipient?.mainClientName || client?.name || 'Client',
    FEEDER_NAME: client?.subTitle || 'GETCO',
    MAIN_METER_NO:
      client?.abtMainMeter?.meterNumber ||
      recipient?.meterNumber ||
      'N/A',
    CHECK_METER_NO: client?.abtCheckMeter?.meterNumber || 'N/A',
    YEAR: period?.year ? String(period.year) : String(new Date().getFullYear()),
    MONTH_NAME: '',
    WEEK_NO: '',
    WEEK_START_DATE: '',
    WEEK_END_DATE: '',
  };

  if (period?.month) {
    const idx = parseInt(period.month, 10) - 1;
    vars.MONTH_NAME = MONTH_NAMES[idx] || String(period.month);
  }

  if (period?.week && period?.year) {
    vars.WEEK_NO = String(period.week);
    const range = getWeekDateRange(period.year, period.week);
    vars.WEEK_START_DATE = range.start;
    vars.WEEK_END_DATE = range.end;
  }

  if (sendType === 'general' && !vars.MONTH_NAME && period?.month) {
    const idx = parseInt(period.month, 10) - 1;
    vars.MONTH_NAME = MONTH_NAMES[idx] || '';
  }

  return vars;
};

const resolveEmailContent = (subjectTemplate, bodyTemplate, variables) => ({
  subject: applyEmailTemplate(subjectTemplate, variables),
  body: applyEmailTemplate(bodyTemplate, variables),
});

module.exports = {
  MONTH_NAMES,
  getWeekDateRange,
  getDefaultSignatureHtml,
  getDefaultTemplates,
  applyEmailTemplate,
  buildEmailVariables,
  resolveEmailContent,
};
