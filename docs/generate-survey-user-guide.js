// Generates docs/Acumatica-Survey-User-Guide.docx
//
// Usage:
//   node docs/generate-survey-user-guide.js
//
// Requires `docx` npm package (`npm install -g docx`).

const fs = require("fs");
const path = require("path");
const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
    Header, Footer, AlignmentType, LevelFormat, HeadingLevel, BorderStyle,
    WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak,
    TableOfContents, PageOrientation, TabStopType, TabStopPosition,
} = require("docx");

// ------------------------------------------------------------------
// Paths
// ------------------------------------------------------------------
const REPO = path.resolve(__dirname, "..");
const IMG = path.join(REPO, "docs", "images");
const OUT = path.join(REPO, "docs", "Acumatica-Survey-User-Guide.docx");

const img = (name) => path.join(IMG, name);

// ------------------------------------------------------------------
// Re-usable helpers
// ------------------------------------------------------------------
const PAGE_W = 12240;            // US Letter width (DXA)
const PAGE_H = 15840;            // US Letter height (DXA)
const MARGIN = 1440;             // 1 inch
const CONTENT_W = PAGE_W - 2 * MARGIN;  // 9360 DXA

function p(text, opts = {}) {
    return new Paragraph({
        children: [new TextRun({ text, ...opts })],
        spacing: { after: 120 },
        ...opts.paragraph,
    });
}

function h1(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun(text)],
        pageBreakBefore: true,
    });
}

function h2(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun(text)],
    });
}

function h3(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun(text)],
    });
}

function bullet(text, level = 0, runs = null) {
    return new Paragraph({
        numbering: { reference: "bullets", level },
        children: runs || [new TextRun({ text })],
        spacing: { after: 80 },
    });
}

function numbered(text, level = 0, runs = null) {
    return new Paragraph({
        numbering: { reference: "steps", level },
        children: runs || [new TextRun({ text })],
        spacing: { after: 80 },
    });
}

function body(text, opts = {}) {
    return new Paragraph({
        children: [new TextRun({ text, ...opts })],
        spacing: { after: 160 },
        alignment: AlignmentType.LEFT,
    });
}

function bodyRuns(runs) {
    return new Paragraph({
        children: runs,
        spacing: { after: 160 },
        alignment: AlignmentType.LEFT,
    });
}

// Image paragraph - tolerates missing file by emitting placeholder
function screenshot(fileName, widthPx = 560, alt = "Acumatica screenshot") {
    const filePath = img(fileName);
    if (!fs.existsSync(filePath)) {
        return new Paragraph({
            children: [new TextRun({ text: `[missing screenshot: ${fileName}]`, italics: true, color: "AA0000" })],
            spacing: { after: 160 },
        });
    }
    const data = fs.readFileSync(filePath);

    // Approximate aspect ratio from file size heuristic - keep a reasonable height.
    // Users can resize in Word. We use 2:1 (w/h) as a safe default for Acumatica
    // screenshots captured at modest aspect, but figures auto-scale in the layout.
    const heightPx = Math.round(widthPx * 0.55);

    const lower = fileName.toLowerCase();
    let type = "png";
    if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) type = "jpg";
    else if (lower.endsWith(".gif")) type = "gif";

    return new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new ImageRun({
            type,
            data,
            transformation: { width: widthPx, height: heightPx },
            altText: { title: fileName, description: alt, name: fileName },
        })],
        spacing: { after: 200 },
    });
}

function caption(text) {
    return new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: `Figure: ${text}`, italics: true, size: 20, color: "555555" })],
        spacing: { after: 280 },
    });
}

function border(color = "CCCCCC", size = 1) {
    return { style: BorderStyle.SINGLE, size, color };
}
function allBorders(color = "CCCCCC") {
    const b = border(color);
    return { top: b, bottom: b, left: b, right: b };
}

// Simple info table with header row and body rows
function infoTable(headers, rows) {
    const colCount = headers.length;
    const colWidth = Math.floor(CONTENT_W / colCount);
    const columnWidths = Array(colCount).fill(colWidth);
    // Fix rounding so the sum matches exactly
    columnWidths[columnWidths.length - 1] = CONTENT_W - colWidth * (colCount - 1);

    const headerCells = headers.map((h, i) => new TableCell({
        borders: allBorders(),
        width: { size: columnWidths[i], type: WidthType.DXA },
        shading: { fill: "2E75B6", type: ShadingType.CLEAR, color: "auto" },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
            children: [new TextRun({ text: h, bold: true, color: "FFFFFF" })],
        })],
    }));

    const bodyRows = rows.map(rowCells => new TableRow({
        children: rowCells.map((cell, i) => new TableCell({
            borders: allBorders(),
            width: { size: columnWidths[i], type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: Array.isArray(cell)
                ? cell
                : [new Paragraph({ children: [new TextRun({ text: String(cell) })] })],
        })),
    }));

    return new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths,
        rows: [new TableRow({ tableHeader: true, children: headerCells }), ...bodyRows],
    });
}

// Tip/Note callout block
function callout(kind, text) {
    const palette = {
        Note: { bg: "E7F1FB", color: "2E75B6" },
        Tip: { bg: "E8F5E9", color: "2E7D32" },
        Warning: { bg: "FFF4E5", color: "B26A00" },
    }[kind] || { bg: "F2F2F2", color: "444444" };

    return new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [CONTENT_W],
        rows: [new TableRow({
            children: [new TableCell({
                borders: {
                    top: { style: BorderStyle.SINGLE, size: 6, color: palette.color },
                    bottom: { style: BorderStyle.SINGLE, size: 6, color: palette.color },
                    left: { style: BorderStyle.SINGLE, size: 24, color: palette.color },
                    right: { style: BorderStyle.SINGLE, size: 6, color: palette.color },
                },
                width: { size: CONTENT_W, type: WidthType.DXA },
                shading: { fill: palette.bg, type: ShadingType.CLEAR, color: "auto" },
                margins: { top: 120, bottom: 120, left: 200, right: 160 },
                children: [new Paragraph({
                    children: [
                        new TextRun({ text: `${kind}: `, bold: true, color: palette.color }),
                        new TextRun({ text }),
                    ],
                })],
            })],
        })],
    });
}

// ------------------------------------------------------------------
// Document sections
// ------------------------------------------------------------------

function coverPage() {
    return [
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 2400, after: 400 },
            children: [new TextRun({
                text: "Acumatica Surveys",
                size: 72, bold: true, color: "2E75B6",
            })],
        }),
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
            children: [new TextRun({
                text: "User Guide",
                size: 44, color: "555555",
            })],
        }),
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 1200 },
            children: [new TextRun({
                text: "Configure, send and track surveys from inside Acumatica ERP",
                italics: true, size: 26, color: "666666",
            })],
        }),
        screenshot("Survey-Workspace.PNG", 540, "Survey workspace in Acumatica"),
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 600 },
            children: [
                new TextRun({ text: "Customization package: ", color: "555555" }),
                new TextRun({ text: "Survey26R1.zip", bold: true }),
            ],
        }),
        new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
                new TextRun({ text: "Acumatica version: ", color: "555555" }),
                new TextRun({ text: "2026 R1 (26.100)", bold: true }),
            ],
        }),
    ];
}

function tocSection() {
    return [
        h1("Table of Contents"),
        new Paragraph({
            children: [new TableOfContents("Table of Contents", {
                hyperlink: true,
                headingStyleRange: "1-3",
            })],
        }),
        body("To refresh the table of contents in Word, right-click inside it and choose " +
            "“Update Field”, then “Update entire table”."),
    ];
}

function overviewChapter() {
    return [
        h1("1. Overview"),
        body("Acumatica Surveys is a customization package that adds a complete survey / " +
             "questionnaire engine to any Acumatica ERP instance. Originally built as a " +
             "COVID-19 employee wellness tracker, it has grown into a general-purpose tool " +
             "for sending surveys to employees, customers or contacts and capturing their " +
             "responses back inside Acumatica."),

        h2("1.1 What you get after installing the package"),
        body("Once Survey26R1 is imported and published, a new Surveys workspace appears in " +
             "the left-hand menu. It groups together the screens listed below."),
        infoTable(
            ["Screen ID", "Name", "Purpose"],
            [
                ["SU101000", "Survey Preferences",
                    "Numbering sequence, default components, webhook binding, anonymous contact."],
                ["SU204003", "Survey Components",
                    "Reusable HTML/Scriban fragments — welcome page, radio-buttons, progress bar, etc."],
                ["SU201000", "Surveys",
                    "The survey itself: compose the sequence of components and questions, manage recipients, watch collectors and answers."],
                ["SU301000", "Survey Collector",
                    "One record per recipient (the 'ballot'): captures token, status, sent / expired dates and the raw answers."],
                ["SU501000", "Process Surveys",
                    "Mass-generate or mass-send collectors to recipients and start or close surveys."],
            ]
        ),
        body(" "),
        screenshot("Survey-Workspace.PNG", 520, "Surveys workspace on the left-hand menu"),
        caption("Surveys workspace exposed by the customization"),

        h2("1.2 Architecture at a glance"),
        body("End-to-end, the flow looks like this:"),
        bullet("A survey is built in Acumatica from reusable components and stored in the Survey table."),
        bullet("When you start or send a survey, the Process Surveys screen creates a SurveyCollector row for each recipient. Each collector carries a unique token and, for logged-in users, an owner contact."),
        bullet("The recipient opens the survey via a webhook URL. No Acumatica login is required when the survey allows anonymous answers."),
        bullet("Each page submission is persisted into SurveyCollectorData (raw payload) and, once processed, split into SurveyAnswer rows for reporting."),
        bullet("Dashboards and Generic Inquiries aggregate the answers for management."),
    ];
}

function installationChapter() {
    return [
        h1("2. Installation"),
        body("Before you can configure or use surveys, the customization package must be " +
             "imported and published."),

        h2("2.1 Prerequisites"),
        bullet("A running Acumatica ERP 2026 R1 instance (build 26.100.x or later)."),
        bullet("Administrator access (the admin user is used in examples below)."),
        bullet("The customization package Survey26R1.zip included with this release."),

        h2("2.2 Import and publish the package"),
        numbered("Open the Customization Projects screen (SM204505): type “Customization” in the search box and pick Customization Projects."),
        screenshot("SS0-CustomizationProjects.PNG", 520, "Open Customization Projects screen"),
        caption("Opening the Customization Projects screen"),

        numbered("Click Import → Import New Project and choose Survey26R1.zip."),
        screenshot("SS2-ImportNewProject.PNG", 520, "Import New Project menu"),
        caption("Import New Project menu on the Customization Projects screen"),

        numbered("Select the imported row and click Publish. Wait for the status to become Published."),
        screenshot("SS3-PublishPackage.PNG", 520, "Publish the package"),
        caption("Publishing the imported customization project"),

        callout("Note", "After the first publish you may see a prompt asking to overwrite " +
            "existing notification templates or sample survey records. Confirm the overwrite — " +
            "the package ships with a demo “Case Survey” you will use later to verify the flow."),
    ];
}

function configurationChapter() {
    return [
        h1("3. Configuring the Survey Module"),

        h2("3.1 Survey Preferences (SU101000)"),
        body("Survey Preferences is the system-wide configuration screen. You must complete " +
             "at least the Numbering Sequence and Contacts before sending any survey."),
        screenshot("S001 - Survey Preferences.PNG", 540, "Survey Preferences screen"),
        caption("Survey Preferences screen (SU101000) in Modern UI"),

        h3("3.1.1 Numbering Sequences"),
        body("Create or pick a numbering sequence that will generate Survey IDs (for example SV10001, SV10002, …). " +
             "You can create a new sequence directly from the selector by clicking the edit icon."),
        screenshot("SS4-SurveyPreferencesNumberingID.PNG", 520, "Numbering Sequence dropdown"),
        caption("Selecting the Survey Numbering ID"),

        h3("3.1.2 Default components"),
        body("These pointers tell the survey generator which reusable components to use by default when it builds a survey page:"),
        infoTable(
            ["Field", "Purpose"],
            [
                ["Template ID", "Wrapper HTML used for every page of the survey."],
                ["Default Header", "Header shown at the top of every survey page (logo, title)."],
                ["Default Question", "Template for a rendered question row."],
                ["Default Comment", "Template for free-text comment fields."],
                ["Default Footer", "Footer shown at the bottom of every survey page."],
                ["Bad Request Page", "Component rendered when a respondent uses an invalid token."],
            ]
        ),
        body(" "),
        callout("Tip", "Start from the pre-loaded components (SUMAIN, SUWELCOME, SUPROGRESS, " +
            "SUPREVNEXT, SURADIOBUTTONS, SUTEXTAREA, SUTHANKYOU, SUBADREQUEST). They are ready " +
            "to use and cover most scenarios."),

        h3("3.1.3 Notifications and contacts"),
        infoTable(
            ["Field", "Purpose"],
            [
                ["WebHook", "Webhook handler that serves survey pages (Survey)."],
                ["New Notification", "Email template sent when a collector is generated."],
                ["Reminder Notification", "Email template sent when a reminder is triggered."],
                ["Contact", "Default From: contact for all notifications."],
                ["Anonymous Contact", "Contact placeholder used for anonymous collectors."],
            ]
        ),
        body(" "),
        callout("Warning", "Anonymous Contact is required for public (AllowAnonymous=true) surveys. " +
            "Without it the webhook endpoint returns 500 for every anonymous request."),

        h2("3.2 Survey Components (SU204003)"),
        body("Survey Components are reusable HTML / Scriban fragments. They are composed into " +
             "survey pages at generation time. Every component has:"),
        bullet("Component ID — e.g. SURADIOBUTTONS."),
        bullet("Component Type — PH (page header), PF (page footer), HE (header), FO (footer), QU (question), CM (comment), MN (main container)."),
        bullet("Description — shown in selectors and in the page designer."),
        bullet("Body — the Scriban / HTML template that renders the component."),
        body(" "),
        screenshot("SS01 - Surveys Components.PNG", 540, "Survey Components screen"),
        caption("Survey Components screen (SU204003) with the new rich text editor"),

        callout("Note", "The Body field now uses a WYSIWYG rich text editor with a built-in " +
            "Visual / HTML / Plain Text mode switch. Toggle to HTML mode when you need to edit " +
            "the raw Scriban tags such as {{survey.title}} or {{page.next_label}}."),

        h2("3.3 Attributes used as questions"),
        body("Each question in a survey is backed by an Acumatica Attribute (CSAttribute). " +
             "Because attributes are standard Acumatica objects, you can reuse them across " +
             "surveys and even against other entities (Cases, Contacts, …)."),
        screenshot("SS5-CreateAttributesQuestions.PNG", 540, "Define an attribute"),
        caption("Defining an attribute that will serve as a survey question"),
        screenshot("SS5-CreateAttributesQuestions2.PNG", 540, "Attribute values"),
        caption("Adding possible answers to an attribute"),
    ];
}

function usingChapter() {
    return [
        h1("4. Creating a Survey"),

        h2("4.1 Compose the survey"),
        numbered("Open Surveys (SU201000). Click + to start a new survey."),
        numbered("Give it a Title, set the Target (Internal / External / Mobile) and Layout."),
        numbered("Pick a Template (default one comes from Survey Preferences)."),
        numbered("On the Details tab, add the sequence of components (header, questions, comments, footers). Each row has a Component Type and a Component ID."),
        numbered("For question rows, pick an Attribute — its predefined values will be rendered as radio buttons or a drop-down."),
        body(" "),
        screenshot("DemoSurveyCreate-S01.PNG", 540, "Demo survey being composed"),
        caption("Composing a survey from reusable components"),

        callout("Tip", "On Survey Preferences there is a Create Demo Survey button. It generates " +
            "a ready-made satisfaction survey you can use as a reference or a starting point."),

        h2("4.2 Preview the survey"),
        body("The Surveys toolbar provides a Preview / Generate HTML action that lets you render " +
             "the whole survey without sending it. Use it to verify that every question displays " +
             "correctly and that the progress bar advances as expected."),
        screenshot("DemoSurveyCreate-S002.PNG", 540, "Preview the generated survey"),
        caption("Preview and Generate HTML actions"),

        h2("4.3 Pick recipients"),
        body("On the Recipients tab of the Survey screen you can add users or contacts that will " +
             "receive the survey. For internal surveys, users are linked to Employee records; for " +
             "external surveys, Contact records are used."),
        screenshot("SS5-RecipentsSelect.PNG", 540, "Select recipients dialog"),
        caption("Selecting recipients"),
        screenshot("SS5-RecipentsSelect2.PNG", 540, "Confirmed recipients"),
        caption("Confirmed list of recipients on the survey"),
    ];
}

function sendingChapter() {
    return [
        h1("5. Sending Surveys"),

        h2("5.1 Start the survey from the Survey screen"),
        body("The simplest way to send a single survey to its configured recipients is to click " +
             "the Start Survey action on the Survey (SU201000) screen. Acumatica will:"),
        bullet("Generate one SurveyCollector row per recipient, each with a unique token."),
        bullet("Set the survey status to Active."),
        bullet("Fire the New Notification email if Email is one of the recipient channels."),
        body(" "),
        screenshot("SS4SurveyTransacationsSurvey.PNG", 540, "Start survey from toolbar"),
        caption("Start Survey action on the survey screen"),

        h2("5.2 Mass operations from Process Surveys (SU501000)"),
        body("When you need to run a survey against many recipients at once, use Process Surveys. " +
             "The Action dropdown controls what happens:"),
        infoTable(
            ["Action", "What happens"],
            [
                ["Send New Notification",   "Create collectors (if missing) and send the New Notification email."],
                ["Send Reminder Notification", "Send the Reminder Notification email to collectors that are still open."],
                ["Process Answers",         "Convert raw payloads in SurveyCollectorData into SurveyAnswer rows for reporting."],
                ["Close Expired",           "Mark collectors older than Expire After as closed."],
            ]
        ),
        body(" "),
        screenshot("SS5-ProcessSendSurvey.PNG", 540, "Process Surveys screen"),
        caption("The Process Surveys screen (SU501000)"),
        screenshot("ProcessSurveyActionDropDown-SS02.PNG", 520, "Action dropdown"),
        caption("Action dropdown on Process Surveys"),
        screenshot("SS5-SendSurveySend.PNG", 520, "Process executed"),
        caption("After Process — selected rows have been sent"),

        h2("5.3 The webhook URL given to the respondent"),
        body("Each collector carries a token. Emails generated from the notification template contain " +
             "a personalised URL of the form:"),
        bodyRuns([
            new TextRun({
                text: "http://<your-host>/<instance>/Webhooks/<tenant>/<survey-webhook-id>?CollectorToken=<token>&PageNbr=1",
                font: "Consolas", size: 20,
            }),
        ]),
        body("This URL is served by the Survey webhook (PX.Survey.Ext.WebHook.SurveyWebhookServerHandler). " +
             "It renders the correct page of the correct survey for the token, and accepts POSTs as the " +
             "respondent advances from page to page."),

        h2("5.4 Scheduling surveys and business events"),
        body("Two built-in Acumatica mechanisms let you run surveys without any manual action:"),
        bullet("Automation Schedules — run Process Surveys on a cron-like schedule."),
        bullet("Business Events — trigger surveys in response to something happening inside Acumatica (for example, a Case being closed)."),
        body(" "),
        screenshot("SS1-AutomationSchedules.PNG", 520, "Automation Schedules list"),
        caption("Automation Schedules list"),
        screenshot("SS1-BusinessEventsSearch.PNG", 520, "Business Events"),
        caption("Searching for Business Events"),
    ];
}

function respondentChapter() {
    return [
        h1("6. Answering a Survey (Recipient View)"),
        body("Recipients receive an email with the personalised URL. They do not need an Acumatica " +
             "login when the survey allows anonymous answers. The form works equally well on desktop " +
             "and mobile:"),

        screenshot("MobileSS-0.jpeg", 300, "Welcome page on mobile"),
        caption("Welcome page on a mobile device"),
        screenshot("MobileSS-1.jpeg", 300, "First question"),
        caption("A radio-button question"),
        screenshot("MobileSS-2.jpeg", 300, "Mid survey"),
        caption("Progress bar and comments"),
        screenshot("MobileSS-3.png", 300, "Thank-you page"),
        caption("Thank-you page after submission"),

        callout("Note", "Every page submission is persisted immediately into SurveyCollectorData. " +
            "Even if the respondent abandons the survey mid-way, the partial answers are retained " +
            "and can be reviewed later."),
    ];
}

function trackingChapter() {
    return [
        h1("7. Tracking Results"),

        h2("7.1 Per-collector view (SU301000)"),
        body("The Survey Collector screen shows one respondent's submission in detail:"),
        bullet("Header — token, status, sent-on / expiration dates, identity of the recipient."),
        bullet("Collected Answers tab — the raw payloads captured by the webhook, one row per page submission."),
        body(" "),
        screenshot("Survey-Response-View.PNG", 540, "Survey Response view"),
        caption("Collector (SU301000) showing a respondent's submission"),

        h2("7.2 Aggregated reporting"),
        body("The customization ships with a dashboard that rolls up answers across all collectors " +
             "of a given survey. It is also a good starting point for building your own reports or " +
             "Generic Inquiries."),
        screenshot("Survey-View-Dashboard.PNG", 540, "Survey dashboard"),
        caption("Aggregated Survey dashboard"),
        screenshot("DashboardSS-1.PNG", 540, "Dashboard detail"),
        caption("Dashboard with breakdown of answers"),
        screenshot("MobileDashboardSS-1.jpg", 300, "Dashboard on mobile"),
        caption("The same dashboard on a mobile device"),

        h2("7.3 Answer lifecycle"),
        body("Answers move through the following states:"),
        infoTable(
            ["State", "Meaning"],
            [
                ["Pending (P)",   "Page submitted but Process Answers has not yet turned the raw payload into SurveyAnswer rows."],
                ["Processed (R)", "Payload has been parsed; one SurveyAnswer row exists per question on the submitted page."],
                ["Closed (C)",    "Collector reached the thank-you page — no more submissions are accepted."],
            ]
        ),
        body(" "),
        callout("Tip", "Schedule Process Surveys → Process Answers to run every few minutes. That " +
            "way dashboards always show results in near real time."),
    ];
}

function troubleshootingChapter() {
    return [
        h1("8. Troubleshooting"),
        infoTable(
            ["Symptom", "Likely cause and fix"],
            [
                ["The Surveys workspace does not appear after publish.",
                    "The customization project shows the Screens section empty. Re-import Survey26R1.zip — the package now registers each page as a <Page> entry so the editor's Screens tab is populated."],
                ["Error “Numbering ID is null” when saving a new survey.",
                    "Survey Preferences is missing a Numbering Sequence, or the sequence's CompanyMask does not match the current tenant. Open SU101000 and pick a numbering sequence; save."],
                ["The anonymous survey URL returns HTTP 500.",
                    "The Anonymous Contact field on Survey Preferences is empty. Assign any existing contact (Employee is fine)."],
                ["The respondent sees a blank page with just Save / Cancel.",
                    "Browser cache — press Ctrl + Shift + R. If it persists after a fresh publish, check D:\\Instances\\<ver>\\<site>\\App_Data\\TSScreenInfo\\<tenant>\\SU*.json: the “views” array must not be empty."],
                ["Modern UI toolbar buttons look chaotic.",
                    "That is the default when many graph actions are public. Use @actionConfig decorators in the screen .ts file to hide or group them."],
            ]
        ),
    ];
}

function appendixChapter() {
    return [
        h1("Appendix A. Screen Reference"),
        infoTable(
            ["Screen ID", "Graph",                              "DAC"],
            [
                ["SU101000",  "PX.Survey.Ext.SurveySetupMaint",         "SurveySetup"],
                ["SU201000",  "PX.Survey.Ext.SurveyMaint",              "Survey / SurveyDetail / SurveyUser / SurveyCollector / SurveyAnswer"],
                ["SU204003",  "PX.Survey.Ext.SurveyComponentMaint",     "SurveyComponent"],
                ["SU301000",  "PX.Survey.Ext.SurveyCollectorMaint",     "SurveyCollector / SurveyCollectorData"],
                ["SU501000",  "PX.Survey.Ext.SurveyProcess",            "SurveyFilter (filter) / SurveyCollector (documents)"],
            ]
        ),

        h1("Appendix B. Database Tables"),
        infoTable(
            ["Table",                 "Holds"],
            [
                ["SurveySetup",          "One row per tenant. System-wide defaults."],
                ["Survey",               "Survey definitions."],
                ["SurveyDetail",         "Ordered list of components that make up a survey page sequence."],
                ["SurveyComponent",      "Reusable HTML / Scriban fragments."],
                ["SurveyUser",           "Users/Contacts who will receive a given survey (recipients)."],
                ["SurveyCollector",      "One row per respondent per survey (the 'ballot')."],
                ["SurveyCollectorData",  "Raw payload for each page submission."],
                ["SurveyAnswer",         "Parsed answer — one row per question on a processed page."],
                ["SurveySetupEntity",    "Binds default surveys to an existing Acumatica entity (Case, Contact, …)."],
            ]
        ),
    ];
}

// ------------------------------------------------------------------
// Document assembly
// ------------------------------------------------------------------
const doc = new Document({
    creator: "Acumatica Surveys Team",
    title: "Acumatica Surveys — User Guide",
    description: "Configure, send and track surveys with the PX.Survey.Ext customization.",

    styles: {
        default: {
            document: {
                run: { font: "Arial", size: 22 }, // 11pt default body
                paragraph: { spacing: { after: 120 } },
            },
        },
        paragraphStyles: [
            { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
                run:       { size: 44, bold: true, font: "Arial", color: "2E75B6" },
                paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
            { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
                run:       { size: 32, bold: true, font: "Arial", color: "1F4E79" },
                paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 } },
            { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
                run:       { size: 26, bold: true, font: "Arial", color: "2E75B6" },
                paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
        ],
    },

    numbering: {
        config: [
            { reference: "bullets", levels: [
                { level: 0, format: LevelFormat.BULLET, text: "\u2022",
                  alignment: AlignmentType.LEFT,
                  style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
                { level: 1, format: LevelFormat.BULLET, text: "\u25E6",
                  alignment: AlignmentType.LEFT,
                  style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
            ]},
            { reference: "steps", levels: [
                { level: 0, format: LevelFormat.DECIMAL, text: "%1.",
                  alignment: AlignmentType.LEFT,
                  style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
            ]},
        ],
    },

    sections: [{
        properties: {
            page: {
                size:   { width: PAGE_W, height: PAGE_H },
                margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
            },
        },
        headers: {
            default: new Header({ children: [new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [new TextRun({
                    text: "Acumatica Surveys — User Guide",
                    italics: true, size: 18, color: "888888",
                })],
            })] }),
        },
        footers: {
            default: new Footer({ children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                    new TextRun({ text: "Page ", size: 18, color: "888888" }),
                    new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "888888" }),
                    new TextRun({ text: " of ", size: 18, color: "888888" }),
                    new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: "888888" }),
                ],
            })] }),
        },
        children: [
            ...coverPage(),
            ...tocSection(),
            ...overviewChapter(),
            ...installationChapter(),
            ...configurationChapter(),
            ...usingChapter(),
            ...sendingChapter(),
            ...respondentChapter(),
            ...trackingChapter(),
            ...troubleshootingChapter(),
            ...appendixChapter(),
        ],
    }],
});

Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(OUT, buffer);
    console.log(`Wrote ${OUT} (${buffer.length} bytes)`);
}).catch(err => {
    console.error("Error:", err);
    process.exit(1);
});
