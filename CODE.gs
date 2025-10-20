// =====================
// CONFIGURATION
// =====================
const SPREADSHEET_ID = ""; // Form Responses Sheet ID
const PARTICIPATION_TEMPLATE_ID = ""; // Participation template

// =====================
// TRIGGER FUNCTION
// =====================
function onFormSubmitTrigger(e) {
  generateCertificatesForLastRow();
}

// =====================
// MAIN CERTIFICATE GENERATOR
// =====================
function generateCertificatesForLastRow() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Form Responses 1");
  const row = sheet.getLastRow();

  if (row < 2) return;

  // Form submitter info
  const mainEmail = sheet.getRange(row, 2).getValue(); // Email
  const formFillerName = sheet.getRange(row, 4).getValue(); // Name of form filler

  // Read role values
  const data = {
    chiefJudge: sheet.getRange(row, 8).getValue(),
    contestChair: sheet.getRange(row, 9).getValue(),
    saa: sheet.getRange(row, 10).getValue(),
    timer: sheet.getRange(row, 11).getValue(),
    ballotCounter: sheet.getRange(row, 12).getValue(),
    firstPlace: sheet.getRange(row, 13).getValue(),
    secondPlace: sheet.getRange(row, 14).getValue(),
    thirdPlace: sheet.getRange(row, 15).getValue(),
    date: sheet.getRange(row, 6).getValue()
  };

  // Participation names from form (column 18)
  const participantsString = sheet.getRange(row, 18).getValue() || "";
  const participants = participantsString
    .split(/\s*,\s*|\n|\r/)
    .filter(name => name.trim() !== "");

  // Role-taker templates
  const ROLES = [
    { key: "firstPlace", label: "First Place", placeholder: "{{First Place}}", templateId: "" },
    { key: "secondPlace", label: "Second Place", placeholder: "{{Second Place}}", templateId: "" },
    { key: "thirdPlace", label: "Third Place", placeholder: "{{Third Place}}", templateId: "" },
    { key: "saa", label: "Sergeant-at-Arms", placeholder: "{{SAA}}", templateId: "" },
    { key: "timer", label: "Timer", placeholder: "{{Timer}}", templateId: "" },
    { key: "ballotCounter", label: "Ballot Counter", placeholder: "{{Ballot Counter}}", templateId: "" },
    { key: "contestChair", label: "Contest Chair", placeholder: "{{Contest Chair}}", templateId: "" }
  ];

  const attachments = [];
  let htmlList = "";

  // Generate role-taker certificates
  ROLES.forEach(role => {
    const participantName = data[role.key];
    if (participantName) {
      htmlList += `<li><b>${role.label}:</b> ${participantName}</li>`;

      const copy = DriveApp.getFileById(role.templateId).makeCopy();
      const presentation = SlidesApp.openById(copy.getId());

      presentation.getSlides().forEach(slide => {
        slide.replaceAllText(role.placeholder, participantName);
        slide.replaceAllText("{{Chief Judge}}", data.chiefJudge);
        slide.replaceAllText("{{Date}}", data.date);
      });
      presentation.saveAndClose();

      attachments.push(copy.getAs("application/pdf").setName(`${participantName} - ${role.label} Certificate.pdf`));
      DriveApp.getFileById(copy.getId()).setTrashed(true);
    }
  });

  // Generate participation certificates
  participants.forEach(name => {
    const copy = DriveApp.getFileById(PARTICIPATION_TEMPLATE_ID).makeCopy();
    const presentation = SlidesApp.openById(copy.getId());

    presentation.getSlides().forEach(slide => {
      slide.replaceAllText("{{Participant}}", name); // Correct placeholder
      slide.replaceAllText("{{Chief Judge}}", data.chiefJudge);
      slide.replaceAllText("{{Date}}", data.date);
    });
    presentation.saveAndClose();

    attachments.push(copy.getAs("application/pdf").setName(`${name} - Participation Certificate.pdf`));
    DriveApp.getFileById(copy.getId()).setTrashed(true);

    htmlList += `<li><b>Participant:</b> ${name}</li>`;
  });

  // Send one email with all attachments
  if (attachments.length > 0) {
    const htmlBody = `
      <p>Dear <b>${formFillerName}</b>,</p>
      <p>Thank you for using <b>AutoCert D91</b>, developed by the <b>District 91 Toastmasters Team</b>!</p>
      <p>Here are the certificates you requested:</p>
      <ul>
        ${htmlList}
      </ul>
      <p>If you have any feedback regarding this system, please let us know — we are improving it day by day.</p>
      <p>Best regards,<br><b>District 91 Toastmasters Team</b></p>
    `;

    MailApp.sendEmail({
      to: mainEmail,
      subject: "Your Contest Certificates",
      body: `Dear ${formFillerName},\n\nThank you for using AuCert D91, developed by the District 91 Toastmasters Team. Here are your certificates (see attachments).\n\nIf you have any feedback, please let us know.\n\n- District 91 Toasmasters Team`,
      htmlBody: htmlBody,
      attachments: attachments
    });
  }

  Logger.log("All certificates (role-taker + participation) generated and emailed!");
}
