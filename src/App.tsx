import { useState, useEffect, useRef, ChangeEvent } from "react";
import { 
  QrCode, 
  Scan, 
  Settings, 
  CheckCircle2, 
  XCircle, 
  AlertTriangle, 
  Download, 
  Copy, 
  Check, 
  Sparkles, 
  Database, 
  Camera, 
  CameraOff,
  User, 
  Calendar,
  Clock,
  ExternalLink,
  Trash2,
  RefreshCw,
  Eye,
  EyeOff,
  Upload
} from "lucide-react";
import { Html5Qrcode } from "html5-qrcode";
import QRCode from "qrcode";

const GOOGLE_APPS_SCRIPT_CODE = `function doPost(e) {
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        // Safe fallback
      }
    }
    
    var sheetId = data.sheetId || "";
    var tabName = data.tabName || data.sheetName || "";
    var ticket = data.ticket || data.ticketNo || "";
    var eventCode = data.eventCode || "";
    var action = data.action || "checkin";
    
    // Auto-extract Sheet ID if they pasted the entire spreadsheet URL
    if (sheetId.indexOf("http") === 0 || sheetId.indexOf("docs.google.com") !== -1) {
      var matches = sheetId.match(/\\/d\\/([a-zA-Z0-9-_]+)\\//);
      if (!matches) {
        matches = sheetId.match(/\\/d\\/([a-zA-Z0-9-_]+)/);
      }
      if (matches && matches[1]) {
        sheetId = matches[1];
      }
    }
    
    if (!sheetId) {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: false,
        valid: false, 
        status: "ERROR",
        reason: "missing_sheet_id",
        message: "Spreadsheet ID is missing in request. Please configure it in settings first."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "fetch_sheets" || action === "get_sheets" || action === "fetch_tabs") {
      var ssList;
      try {
        ssList = SpreadsheetApp.openById(sheetId);
      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          valid: false,
          status: "ERROR",
          reason: "invalid_sheet_id",
          message: "Unable to open Google Sheet. Please confirm your Master Spreadsheet ID is valid."
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var sheetsList = ssList.getSheets();
      var tabNamesList = [];
      for (var i = 0; i < sheetsList.length; i++) {
        tabNamesList.push(sheetsList[i].getName());
      }
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        sheets: tabNamesList,
        tabs: tabNamesList
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (!tabName) {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: false,
        valid: false, 
        status: "ERROR",
        reason: "missing_tab_name",
        message: "Sheet Tab Name is required for query. QR payload is missing a Tab Name middle segment."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (!ticket) {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: false,
        valid: false, 
        status: "ERROR",
        reason: "missing_ticket",
        message: "Ticket number is missing in the request."
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var ss;
    try {
      ss = SpreadsheetApp.openById(sheetId);
    } catch (ssErr) {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: false,
        valid: false, 
        status: "ERROR",
        reason: "invalid_sheet_id",
        message: "Unable to open Google Sheet. Please confirm your Master Spreadsheet ID is valid and sharing permissions allow access."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: false,
        valid: false, 
        status: "ERROR",
        reason: "tab_not_found",
        message: "Spreadsheet Tab '" + tabName + "' was not found inside the Master Spreadsheet. Please create this tab or check spelling."
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Read range
    var range = sheet.getDataRange();
    var values = range.getValues();
    if (values.length === 0 || (values.length === 1 && values[0].length === 1 && !values[0][0])) {
      var defaultHeaders = ["Ticket No", "Name", "Status", "Event Code", "Check In Time", "Event"];
      sheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
      values = sheet.getDataRange().getValues();
    }
    
    var headers = values[0];
    var ticketCol = -1;
    var statusCol = -1;
    var eventCodeCol = -1;
    var checkInCol = -1;
    var eventCol = -1;
    var nameCol = -1;
    
    // Pass 1: Match strict/strong headers case-insensitively
    for (var i = 0; i < headers.length; i++) {
      if (headers[i] === undefined || headers[i] === null) continue;
      
      var hRaw = headers[i].toString().replace(/[\s\u00A0]+/g, ' ').trim();
      var hClean = hRaw.toUpperCase();
      var hAlpha = hRaw.toLowerCase().replace(/[^a-z0-9]/g, '');
      
      if (hClean === "TICKET NO" || hClean === "TICKET NO." || hClean === "TICKET #" || 
          hAlpha === "ticketno" || hAlpha === "ticketnumber" || hAlpha === "ticketcode" || hAlpha === "ticketid") {
        ticketCol = i;
      } else if (hClean === "STATUS" || hAlpha === "status" || hAlpha === "state" || hAlpha === "checkedin" || hAlpha === "checked-in") {
        statusCol = i;
      } else if (hClean === "EVENT CODE" || hClean === "EVENTCODE" || hAlpha === "eventcode") {
        eventCodeCol = i;
      } else if (hClean === "CHECK IN TIME" || hClean === "CHECKINTIME" || hAlpha === "checkintime" || hAlpha === "checkedintime") {
        checkInCol = i;
      } else if (hClean === "EVENT" || hAlpha === "eventname") {
        eventCol = i;
      } else if (hClean === "NAME" || hClean === "GUEST NAME" || hAlpha === "fullname" || hAlpha === "guestname" || hAlpha === "attendeename") {
        nameCol = i;
      }
    }

    // Pass 2: Fallback to weaker/generic keywords ONLY if the column was not resolved in Pass 1
    for (var i = 0; i < headers.length; i++) {
      if (headers[i] === undefined || headers[i] === null) continue;
      
      var hRaw = headers[i].toString().replace(/[\s\u00A0]+/g, ' ').trim();
      var hClean = hRaw.toUpperCase();
      var hAlpha = hRaw.toLowerCase().replace(/[^a-z0-9]/g, '');
      
      if (ticketCol === -1) {
        if (hAlpha === "ticket" || hAlpha === "barcode" || hAlpha === "code" || hAlpha === "qrcode" || hAlpha === "qr" || hAlpha === "ticketid") {
          ticketCol = i;
        }
      }
      if (statusCol === -1 && i !== ticketCol) {
        if (hAlpha === "used" || hAlpha === "checkin" || hAlpha === "attended" || hAlpha === "present" || hAlpha === "arrived") {
          statusCol = i;
        }
      }
      if (eventCodeCol === -1 && i !== ticketCol) {
        if (hAlpha === "codeid") {
          eventCodeCol = i;
        }
      }
      if (checkInCol === -1 && i !== ticketCol) {
        if (hAlpha === "checkin" || hAlpha === "time" || hAlpha === "timestamp") {
          checkInCol = i;
        }
      }
      if (eventCol === -1 && i !== ticketCol) {
        if (hAlpha === "event") {
          eventCol = i;
        }
      }
      if (nameCol === -1 && i !== ticketCol) {
        if (hAlpha === "name" || hAlpha === "guest" || hAlpha === "nama" || hAlpha === "customer" || hAlpha === "client" || hAlpha === "attendee") {
          nameCol = i;
        }
      }
    }

    // Fallback: search for any column names that resemble or contain "ticket", "code", "no", "id", or default to first column
    if (ticketCol === -1) {
      for (var i = 0; i < headers.length; i++) {
        if (headers[i] === undefined || headers[i] === null) continue;
        var hAlpha2 = headers[i].toString().toLowerCase().replace(/[^a-z0-9]/g, '');
        if (hAlpha2.indexOf("ticket") !== -1 || hAlpha2 === "id" || hAlpha2 === "no" || hAlpha2 === "num" || hAlpha2 === "number" || hAlpha2 === "code") {
          ticketCol = i;
          break;
        }
      }
    }
    if (ticketCol === -1 && headers.length > 0) {
      ticketCol = 0;
    }
    
    // Auto-heal missing required check-in fields
    var nextColIndex = headers.length;
    var headersModified = false;
    
    if (ticketCol === -1) {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: false,
        valid: false, 
        status: "ERROR",
        reason: "missing_ticket_header",
        message: "Spreadsheet Tab '" + tabName + "' is missing any identifier column. Please ensure Row 1 has a ticket column."
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (statusCol === -1) {
      nextColIndex++;
      sheet.getRange(1, nextColIndex).setValue("Status");
      statusCol = nextColIndex - 1;
      headersModified = true;
    }
    
    if (checkInCol === -1) {
      nextColIndex++;
      sheet.getRange(1, nextColIndex).setValue("Check In Time");
      checkInCol = nextColIndex - 1;
      headersModified = true;
    }
    
    if (headersModified) {
      values = sheet.getDataRange().getValues();
    }
    
    var foundRow = -1;
    var partialMatchReason = "";
    
    // Parse ticket code using standard replacements
    var searchTicketClean = ticket.toString().replace(/[\s\u00A0]+/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
    var searchEventCodeClean = eventCode.toString().replace(/[\s\u00A0]+/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
    
    for (var r = 1; r < values.length; r++) {
      if (values[r][ticketCol] === undefined || values[r][ticketCol] === null) continue;
      
      var rowTicketVal = values[r][ticketCol].toString().replace(/[\s\u00A0]+/g, ' ').trim();
      
      // Auto-strip trailing .0 from Google Sheets whole-number formatting (e.g. 12345.0 -> 12345)
      if (rowTicketVal.indexOf(".0") === rowTicketVal.length - 2 && rowTicketVal.length > 2) {
        rowTicketVal = rowTicketVal.substring(0, rowTicketVal.length - 2);
      }
      
      var rowTicketClean = rowTicketVal.replace(/[\s\u00A0]+/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
      
      var matchFound = (rowTicketClean === searchTicketClean);
      
      if (!matchFound && rowTicketVal.toLowerCase() === ticket.toString().trim().toLowerCase()) {
        matchFound = true;
      }
      
      // Representation equivalence for integers with leading zeroes (e.g. "00155" matches "155")
      if (!matchFound) {
        var rowNum = parseInt(rowTicketClean, 10);
        var searchNum = parseInt(searchTicketClean, 10);
        if (!isNaN(rowNum) && !isNaN(searchNum) && rowNum === searchNum && rowTicketClean.match(/^[0-9]+$/) && searchTicketClean.match(/^[0-9]+$/)) {
          matchFound = true;
        }
      }
      
      if (matchFound) {
        var rowEventCodeVal = "";
        if (eventCodeCol !== -1 && values[r][eventCodeCol] !== undefined && values[r][eventCodeCol] !== null) {
          rowEventCodeVal = values[r][eventCodeCol].toString().trim();
        }
        var rowEventCodeClean = rowEventCodeVal.toLowerCase().replace(/[^a-z0-9]/g, '');
        
        if (searchEventCodeClean && rowEventCodeClean && rowEventCodeClean !== searchEventCodeClean) {
          partialMatchReason = "Ticket found inside '" + tabName + "' but belongs to Event Code '" + rowEventCodeVal + "' instead of scanned event '" + eventCode + "'.";
          continue;
        }
        
        foundRow = r;
        break;
      }
    }
    
    if (foundRow === -1) {
      var failMsg = partialMatchReason || ("Ticket Number '" + ticket + "' was not found inside spreadsheet tab '" + tabName + "'. Ensure you are checking the correct gate/tab.");
      return ContentService.createTextOutput(JSON.stringify({ 
        success: false,
        valid: false, 
        status: "INVALID",
        reason: "not_found",
        message: failMsg
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var nameVal = (nameCol !== -1 && values[foundRow][nameCol]) ? values[foundRow][nameCol].toString().trim() : "Guest";
    var eventVal = (eventCol !== -1 && values[foundRow][eventCol]) ? values[foundRow][eventCol].toString().trim() : "";
    var eventCodeVal = (eventCodeCol !== -1 && values[foundRow][eventCodeCol]) ? values[foundRow][eventCodeCol].toString().trim() : eventCode;
    var statusVal = (statusCol !== -1 && values[foundRow][statusCol]) ? values[foundRow][statusCol].toString().trim() : "";
    var checkInTimeVal = (checkInCol !== -1 && values[foundRow][checkInCol]) ? values[foundRow][checkInCol].toString().trim() : "";
    
    var normStatus = statusVal.toLowerCase();
    
    if (normStatus === "used" || normStatus === "checked in") {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: false,
        valid: false, 
        status: "USED",
        reason: "already_used",
        name: nameVal,
        guestName: nameVal,
        ticketNo: ticket,
        ticket: ticket,
        eventCode: eventCodeVal,
        event: eventVal,
        eventName: eventVal,
        tabName: tabName,
        checkInTime: checkInTimeVal || "Already checked in"
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var checkinTime = Utilities.formatDate(
      new Date(),
      ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone(),
      "dd-MMM-yyyy hh:mm a"
    );
    
    if (action === "lookup") {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        valid: true,
        status: "APPROVED",
        name: nameVal, 
        guestName: nameVal,
        ticketNo: ticket,
        ticket: ticket,
        eventCode: eventCodeVal,
        event: eventVal,
        eventName: eventVal,
        tabName: tabName,
        checkInTime: checkInTimeVal ? checkInTimeVal.toString() : ""
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    sheet.getRange(foundRow + 1, statusCol + 1).setValue("Used");
    sheet.getRange(foundRow + 1, checkInCol + 1).setValue(checkinTime);
    
    return ContentService.createTextOutput(JSON.stringify({ 
      success: true, 
      valid: true,
      status: "APPROVED",
      name: nameVal, 
      guestName: nameVal,
      ticketNo: ticket,
      ticket: ticket,
      eventCode: eventCodeVal,
      event: eventVal,
      eventName: eventVal,
      tabName: tabName,
      checkInTime: checkinTime
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 
      success: false,
      valid: false, 
      status: "ERROR",
      reason: "error", 
      message: error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var sheetId = "";
    if (e && e.parameter) {
      sheetId = e.parameter.sheetId || e.parameter.spreadsheetId || "";
    }
    
    if (sheetId.indexOf("http") === 0 || sheetId.indexOf("docs.google.com") !== -1) {
      var matches = sheetId.match(/\\/d\\/([a-zA-Z0-9-_]+)\\//);
      if (!matches) {
        matches = sheetId.match(/\\/d\\/([a-zA-Z0-9-_]+)/);
      }
      if (matches && matches[1]) {
        sheetId = matches[1];
      }
    }
    
    if (!sheetId) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "ONLINE"
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var ss = SpreadsheetApp.openById(sheetId);
    var sheets = ss.getSheets();
    var tabNames = [];
    for (var i = 0; i < sheets.length; i++) {
      tabNames.push(sheets[i].getName());
    }
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      sheets: tabNames,
      tabs: tabNames
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}`;

export default function App() {
  // --- CONFIG / SETTINGS STATE ---
  const [eventName, setEventName] = useState<string>(() => {
    return localStorage.getItem("thoc_event_name") || "THE HOUSE OF CONNECTIONS Inaugural Event";
  });
  const [eventCode, setEventCode] = useState<string>(() => {
    return localStorage.getItem("thoc_event_code") || "STRANGER-2026";
  });
  const [sheetId, setSheetId] = useState<string>(() => {
    return localStorage.getItem("thoc_sheet_id") || "";
  });
  const [scriptUrl, setScriptUrl] = useState<string>(() => {
    return localStorage.getItem("thoc_script_url") || "https://script.google.com/macros/s/AKfycby1rbBYT2vUVUDJChRRwe4lipjLqxur1OQfCCCfKF0uaNT3gl8NlXzIIQhgotxZuenw/exec";
  });
  const [fallbackTabName, setFallbackTabName] = useState<string>(() => {
    return localStorage.getItem("thoc_fallback_tab_name") || "Stranger_People";
  });

  const [constructorTabName, setConstructorTabName] = useState<string>(() => {
    return localStorage.getItem("thoc_fallback_tab_name") || "Stranger_People";
  });

  const [constructorEventCode, setConstructorEventCode] = useState<string>(() => {
    return localStorage.getItem("thoc_event_code") || "STRANGER-2026";
  });

  const [fetchedSheetTabs, setFetchedSheetTabs] = useState<string[]>(() => {
    try {
      const stored = localStorage.getItem("thoc_fetched_sheet_tabs");
      return stored ? JSON.parse(stored) : ["Stranger_People", "Stranger_Creators", "Pickleball_People", "Pickleball_Creators", "Volunteers", "Sponsors"];
    } catch {
      return ["Stranger_People", "Stranger_Creators", "Pickleball_People", "Pickleball_Creators", "Volunteers", "Sponsors"];
    }
  });

  const [isFetchingSheets, setIsFetchingSheets] = useState<boolean>(false);

  // Async sheet tabs loader
  const fetchSheetNamesFromSpreadsheet = async (explicitSheetId?: string) => {
    const targetSheetId = (explicitSheetId || sheetId || "").trim();
    if (!targetSheetId) {
      showToastNotification("Configure and Save Spreadsheet ID first!");
      return;
    }

    setIsFetchingSheets(true);
    showToastNotification("Querying Google Sheet tabs...");

    const targetUrl = scriptUrl || "https://script.google.com/macros/s/AKfycby1rbBYT2vUVUDJChRRwe4lipjLqxur1OQfCCCfKF0uaNT3gl8NlXzIIQhgotxZuenw/exec";
    
    try {
      const response = await fetch(targetUrl, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain"
        },
        body: JSON.stringify({
          sheetId: targetSheetId,
          action: "fetch_sheets"
        })
      });

      if (!response.ok) {
        throw new Error(`HTTP error code ${response.status}`);
      }

      const txt = await response.text();
      let resData;
      try {
        resData = JSON.parse(txt);
      } catch (e) {
        throw new Error("Unable to parse result payload as JSON. Deploy updated Apps Script first.");
      }

      if (resData && resData.success && (resData.sheets || resData.tabs)) {
        const tabs = resData.sheets || resData.tabs;
        if (Array.isArray(tabs) && tabs.length > 0) {
          setFetchedSheetTabs(tabs);
          localStorage.setItem("thoc_fetched_sheet_tabs", JSON.stringify(tabs));
          showToastNotification(`Successfully fetched ${tabs.length} sheet tabs!`);
          return tabs;
        } else {
          showToastNotification("Google Spreadsheet has no sheet tabs.");
        }
      } else {
        const msg = resData?.message || "Failed to fetch. Confirm Apps Script is updated and Spreadsheet is shared.";
        showToastNotification(msg);
      }
    } catch (err: any) {
      console.error(err);
      showToastNotification(`Fetch Failed: ${err.message || err.toString()}`);
    } finally {
      setIsFetchingSheets(false);
    }
  };

  // Save Configs to localStorage
  const saveConfiguration = (name: string, codeVal: string, sheetIdVal: string, scriptUrlVal: string, fallbackTabVal: string, cooldownVal?: number) => {
    let cleanSheetId = sheetIdVal.trim();
    if (cleanSheetId.includes("docs.google.com")) {
      const match = cleanSheetId.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (match && match[1]) {
        cleanSheetId = match[1];
      }
    }

    localStorage.setItem("thoc_event_name", name);
    localStorage.setItem("thoc_event_code", codeVal);
    localStorage.setItem("thoc_sheet_id", cleanSheetId);
    localStorage.setItem("thoc_script_url", scriptUrlVal.trim());
    localStorage.setItem("thoc_fallback_tab_name", fallbackTabVal.trim());
    if (cooldownVal !== undefined) {
      localStorage.setItem("thoc_cooldown_seconds", String(cooldownVal));
      setCooldownSeconds(cooldownVal);
    }
    
    setEventName(name);
    setEventCode(codeVal);
    setSheetId(cleanSheetId);
    setScriptUrl(scriptUrlVal.trim());
    setFallbackTabName(fallbackTabVal.trim());
    setConstructorTabName(fallbackTabVal.trim());
    setConstructorEventCode(codeVal.trim());
    
    showToastNotification("Configuration Saved Directly");
  };

  // --- ACTIVE VIEW TAB ---
  const [activeTab, setActiveTab] = useState<"scanner" | "generator" | "settings">("scanner");

  // --- CONFIGURABLE COOLDOWN SYSTEM ---
  const [cooldownSeconds, setCooldownSeconds] = useState<number>(() => {
    try {
      const stored = localStorage.getItem("thoc_cooldown_seconds");
      return stored ? parseInt(stored, 10) : 4;
    } catch {
      return 4;
    }
  });

  const [cooldownRemaining, setCooldownRemaining] = useState<number>(0);

  // --- SCANNER LOCKS & TRACKING ---
  const isProcessingRef = useRef<boolean>(false);
  const lastScannedQRRef = useRef<string | null>(null);
  const lastScanTimeRef = useRef<number>(0);
  const wasCameraScanRef = useRef<boolean>(false);

  // --- SCANNER STATES ---
  const [isScanning, setIsScanning] = useState<boolean>(false);
  const [cameraPermissionStatus, setCameraPermissionStatus] = useState<string>("unknown");
  const [scannedResultCode, setScannedResultCode] = useState<string | null>(null);
  
  // Validation status
  const [validationState, setValidationState] = useState<"idle" | "validating" | "success" | "already_used" | "not_found" | "error">("idle");
  const [validationDetails, setValidationDetails] = useState<{
    name?: string;
    ticket?: string;
    eventCode?: string;
    status?: string;
    time?: string;
    reason?: string;
    errorMessage?: string;
  }>({});

  // Cooldown countdown and auto-resume scan listener
  useEffect(() => {
    if (cooldownRemaining <= 0) return;

    const timer = setInterval(() => {
      setCooldownRemaining((prev) => {
        if (prev <= 1) {
          clearInterval(timer);
          // Cooldown finished: Release the processing lock
          isProcessingRef.current = false;
          
          // Auto-resume camera scanning dynamically only if the scan came from the camera
          if (wasCameraScanRef.current && activeTab === "scanner") {
            console.log(`[Cooldown Finished] Auto-resuming scanner frame detection.`);
            setValidationState("idle");
            startCameraScan();
          }
          return 0;
        }
        return prev - 1;
      });
    }, 1000);

    return () => clearInterval(timer);
  }, [cooldownRemaining, activeTab]);

  // Refs for Scanner
  const html5QrCodeRef = useRef<Html5Qrcode | null>(null);
  const qrReaderId = "thoc-qr-reader";

  // --- DEBUG & DETAILED SCAN STATES ---
  const [isDebugMode, setIsDebugMode] = useState<boolean>(() => {
    return localStorage.getItem("thoc_debug_mode") === "true"; // Defaults to false
  });

  const [scanResult, setScanResult] = useState<{
    raw: string;
    eventCode: string;
    ticketCode: string;
    validationText: string;
    statusText: string;
    isValid: boolean;
    isManual: boolean;
  } | null>(null);

  // Future API Preparation function
  const verifyTicket = (evtCode: string, tktCode: string) => {
    console.log({
      eventCode: evtCode,
      ticketCode: tktCode
    });
  };

  // Helper parser for QR format & validation rules
  const parseAndValidateCode = (rawText: string, isManualInput: boolean) => {
    const trimmed = rawText.trim();
    
    if (trimmed.includes("|")) {
      const parts = trimmed.split("|");
      const extEvent = parts[0]?.trim() || "STRANGER-2026";
      const extTicket = parts[1]?.trim() || "";
      verifyTicket(extEvent, extTicket);
      return {
        raw: trimmed,
        eventCode: extEvent,
        ticketCode: extTicket,
        isValid: true,
        validationText: "Splitting QR data structure",
        isManual: isManualInput,
        statusText: "Ready for verification"
      };
    }

    // Default fallback if no | is found
    verifyTicket("STRANGER-2026", trimmed);
    return {
      raw: trimmed,
      eventCode: "STRANGER-2026",
      ticketCode: trimmed,
      isValid: true,
      validationText: "No pipe delimiter detected. Assuming default event.",
      isManual: isManualInput,
      statusText: "Ready for verification"
    };
  };

  // Handle uploaded QR code files using html5-qrcode's scanFile method
  const handleFileChange = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setValidationState("validating");
    setScannedResultCode("Processing QR image file...");

    try {
      // Create isScanning state clean ups
      if (isScanning) {
        await stopCameraScan();
      }

      // Initialize scanner instance on top of qrReaderId container if not present
      let scanner = html5QrCodeRef.current;
      if (!scanner) {
        scanner = new Html5Qrcode(qrReaderId);
        html5QrCodeRef.current = scanner;
      }

      const decodedText = await scanner.scanFile(file, true);
      setScannedResultCode(decodedText);
      
      // Process verification pipeline with false manual scan
      await processTicketVerification(decodedText, false);
      showToastNotification("QR successfully verified from image file!");
    } catch (error: any) {
      console.warn("Image file QR decode issue:", error);
      setValidationState("not_found");
      
      const failedResult = {
        raw: "DECRYPTION_FAILURE",
        eventCode: "N/A",
        ticketCode: "N/A",
        isValid: false,
        validationText: "Could not decode QR from file ❌",
        isManual: false,
        statusText: "Invalid Ticket"
      };
      setScanResult(failedResult);
      setValidationDetails({
        ticket: "No QR Detected",
        eventCode: "N/A",
        reason: "The uploaded image did not have a clear or valid coordinate QR grid. Try another image."
      });
      showToastNotification("Decoding error: No QR detected in file.");
    } finally {
      // Clear file inputs so same file can be scanned sequentially
      e.target.value = "";
    }
  };

  // --- GENERATOR STATES ---
  const [inputTicket, setInputTicket] = useState<string>("THOC-");
  const [ticketNameInput, setTicketNameInput] = useState<string>("");
  const [generatedQrUrl, setGeneratedQrUrl] = useState<string | null>(null);
  const [isGenerating, setIsGenerating] = useState<boolean>(false);
  const [isSyncingName, setIsSyncingName] = useState<boolean>(false);

  const handlePullNameFromSpreadsheet = async (targetCodeOverride?: string, silent: boolean = false) => {
    const code = (typeof targetCodeOverride === "string" ? targetCodeOverride : inputTicket).trim();
    if (!code || code === "THOC-") {
      if (!silent) {
        showToastNotification("Enter a valid ticket code first");
      }
      return;
    }

    setIsSyncingName(true);
    const targetUrl = scriptUrl || "https://script.google.com/macros/s/AKfycby1rbBYT2vUVUDJChRRwe4lipjLqxur1OQfCCCfKF0uaNT3gl8NlXzIIQhgotxZuenw/exec";

    try {
      let response;
      const payload = {
        sheetId: sheetId,
        tabName: constructorTabName,
        sheetName: constructorTabName,
        eventCode: eventCode,
        ticket: code,
        ticketNo: code,
        action: "lookup"
      };

      try {
        response = await fetch(targetUrl, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        });
      } catch (err) {
        response = await fetch(targetUrl, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify(payload)
        });
      }

      if (!response.ok) {
        throw new Error(`HTTP network error status: ${response.status}`);
      }

      const resData = await response.json();
      console.log("Lookup response from spreadsheet:", resData);

      const retrievedName = resData.name || resData.guestName || resData.guest_name || resData.GuestName;
      if ((resData.success === true || resData.valid === true) && retrievedName) {
        setTicketNameInput(retrievedName);
        if (!silent) {
          showToastNotification(`Synced Name: ${retrievedName}`);
        }
      } else {
        if (!silent) {
          if (resData.reason === "not_found") {
            showToastNotification("Ticket details not registered in the spreadsheet");
          } else {
            showToastNotification("Name not found. Make sure ticket is added in sheet tab");
          }
        }
      }
    } catch (err: any) {
      console.warn("Lookup error:", err);
      if (!silent) {
        showToastNotification("Sync failed. Check settings or deployment version");
      }
    } finally {
      setIsSyncingName(false);
    }
  };

  // Toast System
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const [showConfigCode, setShowConfigCode] = useState<boolean>(false);

  const showToastNotification = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => {
      setToastMessage(null);
    }, 3000);
  };

  // --- MANAGE SCANNING LIFE-CYCLE ---
  const startCameraScan = async () => {
    if (isScanning) return;
    
    // Reset scanned values
    setScannedResultCode(null);
    setValidationState("idle");
    setValidationDetails({});

    try {
      setIsScanning(true);
      setCameraPermissionStatus("pending");

      // Verify or initialize html5-qrcode instance
      const scanner = new Html5Qrcode(qrReaderId);
      html5QrCodeRef.current = scanner;

      await scanner.start(
        { facingMode: "environment" },
        {
          fps: 24,
        },
        (decodedText) => {
          // Success Callback
          onSuccessfulQRCodeScan(decodedText);
        },
        () => {
          // Keep silent on scanning frame errors
        }
      );
      setCameraPermissionStatus("granted");
    } catch (err: any) {
      console.warn("Camera access state update:", err);
      setIsScanning(false);
      setCameraPermissionStatus("denied");
      showToastNotification("Camera access was requested but denied.");
    }
  };

  const stopCameraScan = async () => {
    if (!isScanning) return;

    try {
      if (html5QrCodeRef.current && html5QrCodeRef?.current.isScanning) {
        await html5QrCodeRef.current.stop();
        html5QrCodeRef.current = null;
      }
    } catch (err) {
      console.warn("Trouble stopping camera device scan:", err);
    } finally {
      setIsScanning(false);
    }
  };

  // Auto clean camera on navigation
  useEffect(() => {
    if (activeTab !== "scanner") {
      stopCameraScan();
    }
  }, [activeTab]);  // Handle Scan Outcome
  const onSuccessfulQRCodeScan = async (ticketCode: string) => {
    const now = Date.now();
    const cleanCode = ticketCode.trim();

    // 1. Prevent multiple requests while a scan is actively being processed (Processing Lock)
    if (isProcessingRef.current) {
      console.log(`[Scan ignored] QR detected: "${cleanCode}", but another ticket verify request is already active (Processing Lock).`);
      return;
    }

    // 2. Prevent race conditions and identical repeat scans in the cooldown window
    const elapsed = now - lastScanTimeRef.current;
    const cooldownMs = cooldownSeconds * 1000;

    if (cleanCode === lastScannedQRRef.current && elapsed < cooldownMs) {
      const secondsLeft = ((cooldownMs - elapsed) / 1000).toFixed(1);
      console.log(`[Scan ignored due to cooldown] QR detected: "${cleanCode}", but it matches previous scan and cooldown still has ${secondsLeft}s remaining.`);
      return;
    }

    // Capture scanning intent
    console.log(`[QR detected] Decoded string: "${cleanCode}".`);

    // Acquire lock and update trace
    isProcessingRef.current = true;
    wasCameraScanRef.current = true;
    lastScannedQRRef.current = cleanCode;
    lastScanTimeRef.current = now;

    // Vibrate device briefly for tactile feedback if supported
    if ("vibrate" in navigator) {
      try {
        navigator.vibrate(100);
      } catch (e) {}
    }

    // Stop scan immediately after successful code extraction to prevent frame overflow
    console.log("[Camera] Pausing camera lens during API verification and cooldown.");
    await stopCameraScan();
    
    setScannedResultCode(cleanCode);
    await processTicketVerification(cleanCode, false);
  };

  // Process Verification with Google Sheet Apps Script Single Source of Truth API
  const processTicketVerification = async (ticketCode: string, isManualInput: boolean = scanResult?.isManual || false) => {
    const formattedCode = ticketCode.trim();
    if (!formattedCode) return;

    // Ensure we hold the lock on manual inputs or file uploads as well
    isProcessingRef.current = true;
    if (isManualInput) {
      wasCameraScanRef.current = false;
      lastScannedQRRef.current = formattedCode;
      lastScanTimeRef.current = Date.now();
    }

    setValidationState("validating");

    // Split the QR string by "|"
    // EVENT_CODE|TAB_NAME|TICKET_NO
    let eventCodePart = "";
    let tabNamePart = "";
    let ticketNoPart = "";
    
    if (formattedCode.includes("|")) {
      const parts = formattedCode.split("|");
      if (parts.length >= 3) {
        eventCodePart = parts[0]?.trim() || "";
        tabNamePart = parts[1]?.trim() || "";
        ticketNoPart = parts[2]?.trim() || "";
      } else if (parts.length === 2) {
        eventCodePart = parts[0]?.trim() || "";
        tabNamePart = fallbackTabName; // default/fallback sheet tab name from settings
        ticketNoPart = parts[1]?.trim() || "";
      } else {
        eventCodePart = eventCode;
        tabNamePart = fallbackTabName;
        ticketNoPart = parts[0]?.trim() || "";
      }
    } else {
      eventCodePart = eventCode;
      tabNamePart = fallbackTabName;
      ticketNoPart = formattedCode;
    }

    // Capture parsing results locally for transient visual loader validation
    const parsed = {
      raw: formattedCode,
      eventCode: eventCodePart || eventCode || "STRANGER-2026",
      ticketCode: ticketNoPart,
      isValid: true,
      validationText: `Split: Event "${eventCodePart || eventCode}" | Tab "${tabNamePart}" | Ticket "${ticketNoPart}"`,
      isManual: isManualInput,
      statusText: "Validating..."
    };
    setScanResult(parsed);
    setValidationDetails({
      eventCode: eventCodePart || eventCode || "STRANGER-2026",
      ticket: ticketNoPart,
      tabName: tabNamePart
    });

    const targetUrl = scriptUrl || "https://script.google.com/macros/s/AKfycby1rbBYT2vUVUDJChRRwe4lipjLqxur1OQfCCCfKF0uaNT3gl8NlXzIIQhgotxZuenw/exec";

    console.log(`[API request sent] Dispatching check-in request. Ticket: "${ticketNoPart}", Tab: "${tabNamePart}", Event: "${eventCodePart}"`);

    let response;
    try {
      try {
        // Try JSON post as requested by user
        response = await fetch(targetUrl, {
          method: "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            sheetId: sheetId,
            tabName: tabNamePart,
            sheetName: tabNamePart,
            eventCode: eventCodePart,
            ticket: ticketNoPart,
            ticketNo: ticketNoPart,
            action: "checkin"
          })
        });
      } catch (jsonErr) {
        console.warn("JSON fetch failed or CORS preflight blocked. Falling back to simple request with text/plain to bypass OPTIONS preflight:", jsonErr);
        // Fallback to text/plain which acts as a CORS-safe 'simple request' style payload
        response = await fetch(targetUrl, {
          method: "POST",
          headers: {
            "Content-Type": "text/plain;charset=utf-8"
          },
          body: JSON.stringify({
            sheetId: sheetId,
            tabName: tabNamePart,
            sheetName: tabNamePart,
            eventCode: eventCodePart,
            ticket: ticketNoPart,
            ticketNo: ticketNoPart,
            action: "checkin"
          })
        });
      }

      if (!response.ok) {
        throw new Error(`API returned HTTP status connection error: ${response.status}`);
      }

      const resData = await response.json();
      console.log(`[API response received] Payload returned from server.`, resData);

      // Current timestamp fallback format
      const currentTime = new Date().toLocaleTimeString('en-US', {
        hour: '2-digit',
        minute: '2-digit',
        hour12: true
      }) + `, ` + new Date().toLocaleDateString('en-US', {
        day: '2-digit',
        month: 'short',
        year: 'numeric'
      });

      // Extract details flexibly matching typical Apps Script output patterns
      const guestName = resData.name || resData.guestName || resData.guest_name || resData.GuestName || "";
      const respTicketNo = resData.ticketNo || resData.ticket_no || resData.ticketNumber || resData.ticketCode || resData.ticket || ticketNoPart;
      const respEventCode = resData.eventCode || resData.event_code || eventCodePart;
      const respCheckInTime = resData.checkInTime || resData.check_in_time || resData.time || resData.timestamp || currentTime;
      const respTabName = resData.tabName || resData.sheetName || tabNamePart;

      // Handle exact response mappings defined in instructions & user's script
      const isValid = resData.valid === true || resData.success === true;
      const resError = String(resData.error || "").toLowerCase();
      const resReason = String(resData.reason || "").toLowerCase();
      const resStatus = String(resData.status || "").toLowerCase();
      const msgLower = String(resData.message || "").toLowerCase();
      
      const isAlreadyUsed = resReason === "already_used" || 
                           resReason === "already_checked_in" ||
                           resError === "already_used" ||
                           (!isValid && (
                             resStatus === "used" || 
                             resStatus === "checked in" || 
                             resStatus.includes("use") || 
                             resStatus.includes("check") || 
                             resReason.includes("use") ||
                             resReason.includes("check") ||
                             resError.includes("use") ||
                             resError.includes("check") ||
                             msgLower.includes("already") ||
                             msgLower.includes("used") ||
                             msgLower.includes("scanned") ||
                             msgLower.includes("checked")
                           ));

      if (isValid) {
        console.log(`[Scan outcome] APPROVED - Entry granted for ticket "${respTicketNo}". Initiating ${cooldownSeconds}s cooldown.`);
        setValidationState("success");
        setValidationDetails({
          name: guestName || "Guest",
          ticket: respTicketNo,
          eventCode: respEventCode,
          tabName: respTabName,
          time: respCheckInTime,
          status: "APPROVED"
        });
        
        // Start cooldown!
        setCooldownRemaining(cooldownSeconds);
      } else if (isAlreadyUsed) {
        console.log(`[Scan outcome] USED - Ticket "${respTicketNo}" was checked in previously at "${respCheckInTime}". Initiating ${cooldownSeconds}s cooldown.`);
        setValidationState("already_used");
        setValidationDetails({
          name: guestName || "Guest",
          ticket: respTicketNo,
          eventCode: respEventCode,
          tabName: respTabName,
          time: respCheckInTime,
          status: "USED"
        });
        
        // Start cooldown!
        setCooldownRemaining(cooldownSeconds);
      } else {
        console.log(`[Scan outcome] INVALID - Ticket Code "${respTicketNo}" is invalid or failed. Reason: ${resData.reason || resData.message}`);
        setValidationState("not_found");
        setValidationDetails({
          name: guestName || "",
          ticket: respTicketNo,
          eventCode: respEventCode,
          tabName: respTabName,
          time: respCheckInTime,
          status: "INVALID",
          reason: resData.reason || resData.message || "Ticket not found or custom spreadsheet header mismatch."
        });
        
        // Release lock immediately on failed lookup so they can try again / type manually
        isProcessingRef.current = false;
      }
    } catch (err: any) {
      console.warn("Spreadsheet API call encountered an issue:", err);
      setValidationState("error");
      setValidationDetails({
        errorMessage: "CORS / Network Blocked: Make sure your Google Sheet Apps Script is deployed as a Web App, executes as 'Me', has Access set to 'Anyone', and that you are connected to the internet."
      });
      
      // Release lock on exception to prevent being frozen
      isProcessingRef.current = false;
    }
  };

  // Manual code check-in from scanner screen directly
  const handleManualCheckIn = (code: string) => {
    if (!code || !code.trim()) return;
    setScannedResultCode(code.trim());
    processTicketVerification(code.trim(), true);
  };

  // --- TICKET GENERATOR HELPERS ---
  const generateRandomTicketCode = () => {
    const chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
    let randomStr = "";
    for (let i = 0; i < 8; i++) {
      randomStr += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    const prefix = constructorEventCode ? constructorEventCode.split('-')[0] : "THOC";
    setInputTicket(`${prefix}-${randomStr}`);
  };

  const handleGenerateQR = async () => {
    const formatted = inputTicket.trim();
    
    // Pattern to match 7 to 20 characters of letters, numbers, and hyphens
    const pattern = /^[a-zA-Z0-9-]{7,20}$/;
    if (!pattern.test(formatted)) {
      showToastNotification("Ticket must be between 7 and 20 characters (alphanumeric and dashes accepted)");
      return;
    }

    setIsGenerating(true);
    
    try {
      // Encode with standard luxury triple segment format: EVENT_CODE|TAB_NAME|TICKET_NUMBER
      const qrDataString = `${constructorEventCode.trim()}|${constructorTabName.trim()}|${formatted}`;

      const dataUrl = await QRCode.toDataURL(qrDataString, {
        width: 380,
        margin: 2,
        color: {
          dark: "#000000",
          light: "#FFFFFF"
        },
        errorCorrectionLevel: 'H'
      });

      setGeneratedQrUrl(dataUrl);
      showToastNotification("QR Code Rendered Successfully.");
    } catch (err) {
      console.warn(err);
      showToastNotification("Could not render QR code.");
    } finally {
      setIsGenerating(false);
    }
  };

  const downloadQRAsPNG = () => {
    if (!generatedQrUrl) return;
    const a = document.createElement("a");
    a.href = generatedQrUrl;
    a.download = `THOC_Ticket_${inputTicket.trim()}.png`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    showToastNotification("Downloaded PNG Ticket");
  };

  return (
    <div id="thoc-app-container" className="min-h-screen bg-[#070707] flex flex-col items-center justify-start py-4 px-3 sm:px-6 select-none font-sans text-stone-200">
      
      {/* BRANDING LOGO */}
      <header id="thoc-header" className="w-full max-w-md bg-black/40 backdrop-blur-md rounded-2xl mb-4 border border-zinc-900 overflow-hidden">
        <div className="py-4 px-6 text-center">
          <h1 className="font-serif text-xl sm:text-2xl font-bold tracking-[0.25em] text-white">
            THE HOUSE OF CONNECTIONS
          </h1>
          <p className="font-display text-[9px] uppercase tracking-[0.43em] text-zinc-500 mt-1">
            THoC Event Ticket Manager
          </p>
        </div>
        
        {/* CURRENT EVENT NAME SHOWCASE */}
        <div className="bg-zinc-950/90 py-2.5 px-4 text-center border-t border-zinc-900/60 overflow-hidden whitespace-nowrap text-ellipsis flex items-center justify-center gap-2">
          <Sparkles className="w-3.5 h-3.5 text-zinc-400 shrink-0" />
          <span className="font-display text-xs text-zinc-300 font-medium lowercase tracking-wide first-letter:uppercase">
            {eventName}
          </span>
          <span className="text-[9px] font-mono px-1.5 py-0.5 rounded-md border shrink-0 bg-emerald-950/20 text-emerald-500 border-emerald-900/50">
            LIVE SHEETS
          </span>
        </div>
      </header>

      {/* STICKY MAIN CONTENT CONTAINER (Android Phone First Dimensions) */}
      <main id="thoc-main-stage" className="w-full max-w-md flex-1 flex flex-col bg-black border border-zinc-900 rounded-3xl overflow-hidden relative shadow-2xl shadow-black">
        
        {/* VIEW CONTAINER */}
        <div className="flex-1 p-4 overflow-y-auto">
          
          {/* ======================================= */}
          {/* TAB 1: SCANNER SECTION                  */}
          {/* ======================================= */}
          {activeTab === "scanner" && (
            <div id="tab-scanner-section" className="flex flex-col h-full gap-4">
              
              {/* HEADING AND DIRECT ACTION */}
              <div className="flex justify-between items-center pb-2 border-b border-zinc-900">
                <div className="flex items-center gap-2">
                  <Scan className="w-4 h-4 text-zinc-400" />
                  <h2 className="font-display text-sm uppercase tracking-wider font-semibold text-white">
                    Gate Check-In
                  </h2>
                </div>
                <span className="text-[10px] font-mono text-zinc-500">
                  Google Scripts
                </span>
              </div>

              {/* LIVE CAMERA VIEWFRAME */}
              <div className={`relative w-full rounded-2xl border border-zinc-800 bg-zinc-950 overflow-hidden flex flex-col items-center justify-center group ${(validationState === "idle" || validationState === "validating") ? "aspect-square" : ""}`}>
                
                {/* Visual Target Reticle (Always on screen if checking) */}
                {isScanning && (
                  <div className="absolute inset-0 z-10 pointer-events-none flex items-center justify-center">
                    {/* The scanning matrix border bracket indicators */}
                    <div className="w-60 h-60 relative flex items-center justify-center">
                      {/* Corner indicators */}
                      <div className="absolute top-0 left-0 w-6 h-6 border-t-2 border-l-2 border-white rounded-tl-md"></div>
                      <div className="absolute top-0 right-0 w-6 h-6 border-t-2 border-r-2 border-white rounded-tr-md"></div>
                      <div className="absolute bottom-0 left-0 w-6 h-6 border-b-2 border-l-2 border-white rounded-bl-md"></div>
                      <div className="absolute bottom-0 right-0 w-6 h-6 border-b-2 border-r-2 border-white rounded-br-md"></div>
                      
                      {/* Laser sweep line animation */}
                      <div className="w-full h-0.5 bg-gradient-to-r from-transparent via-red-500 to-transparent absolute top-0 left-0 animate-laser shadow-[0_0_8px_rgba(239,68,68,0.8)]"></div>
                    </div>
                  </div>
                )}

                {/* Html5-qrcode mount pointer target */}
                <div 
                  id={qrReaderId} 
                  className={`w-full h-full overflow-hidden transition-opacity duration-300 ${isScanning ? "opacity-100" : "opacity-0 hidden"}`}
                ></div>

                {/* Idle screen state */}
                {!isScanning && validationState === "idle" && (
                  <div className="flex flex-col items-center p-6 text-center pointer-events-none">
                    <div className="w-16 h-16 rounded-full bg-zinc-900/80 border border-zinc-800 flex items-center justify-center mb-3">
                      <Camera className="w-8 h-8 text-zinc-400" />
                    </div>
                    <p className="text-zinc-200 text-sm font-semibold tracking-wide">
                      Camera Lens Static
                    </p>
                    <p className="text-zinc-500 text-xs mt-1.5 max-w-[200px]">
                      Access device camera dynamically below to scan tickets
                    </p>
                  </div>
                )}

                {/* VALIDATING LOADER OVERLAY */}
                {validationState === "validating" && (
                  <div className="absolute inset-0 bg-black/95 z-20 flex flex-col items-center justify-center text-center p-6 rounded-2xl border-2 border-zinc-800">
                    <div className="relative flex items-center justify-center w-14 h-14 mb-4">
                      {/* Premium loading spinner */}
                      <div className="absolute inset-0 rounded-full border-2 border-zinc-800"></div>
                      <div className="absolute inset-0 rounded-full border-t-2 border-white animate-spin"></div>
                    </div>
                    <p className="text-white text-sm font-mono tracking-widest font-semibold uppercase animate-pulse">
                      AUTHENTICATING TICKET...
                    </p>
                    
                    {/* Raw verification value */}
                    <div className="bg-zinc-950 border border-zinc-900 rounded-lg px-2.5 py-1.5 mt-2.5 max-w-[240px] w-full text-center">
                      <span className="text-[10px] text-zinc-500 block uppercase tracking-widest font-bold">Raw input:</span>
                      <span className="text-[11px] text-zinc-300 font-mono break-all font-semibold uppercase tracking-wider block mt-0.5 mt-0.5">
                        {scannedResultCode}
                      </span>
                    </div>

                    {/* Extracted Event & Ticket indicators */}
                    <div className="bg-zinc-950 border border-zinc-900 rounded-lg p-3 mt-2.5 max-w-[240px] w-full text-left font-mono space-y-1.5">
                      <div className="flex justify-between items-baseline">
                        <span className="text-[10px] text-zinc-500 uppercase tracking-wider">EVENT CODE:</span>
                        <span className="text-xs text-amber-500 font-bold tracking-wider uppercase">
                          {scanResult?.eventCode || "Extracting..."}
                        </span>
                      </div>
                      <div className="flex justify-between items-baseline">
                        <span className="text-[10px] text-zinc-500 uppercase tracking-wider">TICKET CODE:</span>
                        <span className="text-xs text-stone-200 font-bold tracking-wider uppercase">
                          {scanResult?.ticketCode || "Extracting..."}
                        </span>
                      </div>
                    </div>
                  </div>
                )}

                {/* VALIDATION SCREEN: MULTIPLE RESPONSIVE STATUS STATES */}
                
                 {/* A. SUCCESS: GREEN CARD */}
                 {validationState === "success" && (
                  <div className="w-full bg-black z-30 flex flex-col p-5 border-2 border-emerald-500 rounded-2xl relative">
                    <div className="flex-1 flex flex-col justify-center items-center text-center">
                      
                      <div className="w-14 h-14 bg-emerald-950/60 border border-emerald-500/40 rounded-full flex items-center justify-center mb-3 animate-bounce shadow-[0_0_15px_rgba(16,185,129,0.3)]">
                        <CheckCircle2 className="w-9 h-9 text-emerald-400" />
                      </div>
                      
                      <span className="font-mono text-xs bg-emerald-950/50 text-emerald-400 border border-emerald-500/20 px-3 py-1 rounded-full font-semibold uppercase tracking-widest leading-none mb-1">
                        CHECK-IN OK
                      </span>

                      <h2 className="font-display text-xl sm:text-2xl font-bold tracking-tight text-white mt-1 uppercase">
                        ✅ ENTRY APPROVED
                      </h2>

                      {/* TICKET DETAILS GRID */}
                      <div className="w-full bg-zinc-950 border border-zinc-900 rounded-xl px-4 py-3.5 mt-4 space-y-2.5 text-left font-mono">
                        <div className="flex justify-between items-baseline border-b border-zinc-900 pb-1.5">
                          <span className="text-[10px] text-zinc-500 uppercase tracking-wider">Event Code</span>
                          <span className="text-xs text-amber-500 font-bold tracking-widest uppercase">
                            {validationDetails.eventCode}
                          </span>
                        </div>

                        <div className="flex justify-between items-baseline border-b border-zinc-900 pb-1.5">
                          <span className="text-[10px] text-zinc-500 uppercase tracking-wider">Guest Name</span>
                          <div className="flex items-center gap-1.5 shrink-0 max-w-[180px]">
                            <span className="text-xs text-stone-200 font-bold tracking-tight truncate">
                              {validationDetails.name || "Guest"}
                            </span>
                            {validationDetails.name && (
                              <span className="text-[8px] bg-emerald-950 text-emerald-400 border border-emerald-500/20 px-1 py-0.5 rounded font-bold uppercase tracking-widest leading-none">
                                Spreadsheet
                              </span>
                            )}
                          </div>
                        </div>
                        
                        <div className="flex justify-between items-baseline border-b border-zinc-900 pb-1.5">
                          <span className="text-[10px] text-zinc-500 uppercase tracking-wider">Ticket Number</span>
                          <span className="text-xs text-white tracking-widest font-bold">
                            {validationDetails.ticket}
                          </span>
                        </div>

                        <div className="flex justify-between items-baseline">
                          <span className="text-[10px] text-zinc-500 uppercase tracking-wider">Checked In At</span>
                          <span className="text-[11px] text-zinc-400">
                            {validationDetails.time}
                          </span>
                        </div>

                        {validationDetails.tabName && (
                          <div className="flex justify-between items-baseline border-t border-zinc-900/60 pt-2 text-left">
                            <span className="text-[10px] text-zinc-500 uppercase tracking-wider">Spreadsheet Tab (Tier)</span>
                            <span className="text-xs text-stone-300 font-semibold uppercase tracking-wider bg-zinc-905 px-2 py-0.5 rounded border border-zinc-850">
                              {validationDetails.tabName}
                            </span>
                          </div>
                        )}
                      </div>

                      {/* Cool Animated Cooldown Indicator Bar */}
                      {cooldownRemaining > 0 && (
                        <div className="w-full mt-4 bg-zinc-950 border border-zinc-900 p-2.5 rounded-xl space-y-1">
                          <div className="flex justify-between items-center text-[9px] font-mono uppercase tracking-wider text-emerald-400">
                            <span>Camera paused - Cooldown active</span>
                            <span className="font-bold">{cooldownRemaining}s remaining</span>
                          </div>
                          <div className="w-full bg-zinc-900 h-1.5 rounded-full overflow-hidden">
                            <div 
                              className="bg-emerald-400 h-full rounded-full transition-all duration-300 ease-out shadow-[0_0_8px_rgba(52,211,153,0.5)]"
                              style={{ width: `${(cooldownRemaining / cooldownSeconds) * 100}%` }}
                            ></div>
                          </div>
                        </div>
                      )}

                    </div>

                    <button 
                      onClick={() => {
                        setValidationState("idle"); 
                        startCameraScan();
                      }}
                      className="w-full bg-emerald-600 hover:bg-emerald-500 text-white font-mono text-xs tracking-widest uppercase py-3 rounded-xl font-bold cursor-pointer transition-all shrink-0 shadow-lg shadow-emerald-900/30 flex justify-center items-center gap-2"
                    >
                      {cooldownRemaining > 0 ? (
                        <>
                          <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                          <span>AUTO-RESUMING IN {cooldownRemaining}S</span>
                        </>
                      ) : (
                        "DISMISS & SCAN NEXT"
                      )}
                    </button>
                  </div>
                )}

                 {/* B. BLOCKED/ALREADY USED: RED CARD */}
                 {validationState === "already_used" && (
                  <div className="w-full bg-black z-30 flex flex-col p-5 border-2 border-red-500 rounded-2xl relative">
                    <div className="flex-1 flex flex-col justify-center items-center text-center">
                      
                      <div className="w-14 h-14 bg-red-950/60 border border-red-500/40 rounded-full flex items-center justify-center mb-3 animate-pulse shadow-[0_0_15px_rgba(239,68,68,0.3)]">
                        <XCircle className="w-9 h-9 text-red-500" />
                      </div>
                      
                      <span className="font-mono text-xs bg-red-950/50 text-red-400 border border-red-500/20 px-3 py-1 rounded-full font-semibold uppercase tracking-widest leading-none mb-1">
                        DUPLICATE CODE
                      </span>

                      <h2 className="font-display text-lg sm:text-xl font-bold tracking-tight text-white mt-1 uppercase">
                        ❌ TICKET ALREADY USED
                      </h2>

                      {/* TICKET DETAILS GRID */}
                      <div className="w-full bg-zinc-950 border border-zinc-900 rounded-xl px-4 py-3 mt-4 space-y-2.5 text-left font-mono">
                        <div className="flex justify-between items-baseline border-b border-zinc-900 pb-1.5">
                          <span className="text-[10px] text-zinc-500 uppercase tracking-wider">Event Code</span>
                          <span className="text-xs text-red-400 font-bold tracking-widest uppercase">
                            {validationDetails.eventCode}
                          </span>
                        </div>

                        <div className="flex justify-between items-baseline border-b border-zinc-900 pb-1.5">
                          <span className="text-[10px] text-zinc-500 uppercase tracking-wider">Guest Name</span>
                          <div className="flex items-center gap-1.5 shrink-0 max-w-[180px]">
                            <span className="text-xs text-stone-200 font-bold tracking-tight truncate">
                              {validationDetails.name || "Guest"}
                            </span>
                            {validationDetails.name && (
                              <span className="text-[8px] bg-red-950 text-red-400 border border-red-500/20 px-1 py-0.5 rounded font-bold uppercase tracking-widest leading-none">
                                Spreadsheet
                              </span>
                            )}
                          </div>
                        </div>
                        
                        <div className="flex justify-between items-baseline border-b border-zinc-900 pb-1.5">
                          <span className="text-[10px] text-zinc-500 uppercase tracking-wider">Code ID</span>
                          <span className="text-xs text-zinc-400 tracking-widest">
                            {validationDetails.ticket}
                          </span>
                        </div>

                        {validationDetails.tabName && (
                          <div className="flex justify-between items-baseline border-b border-zinc-900 pb-1.5 text-left">
                            <span className="text-[10px] text-zinc-500 uppercase tracking-wider">Spreadsheet Tab (Tier)</span>
                            <span className="text-xs text-red-300 font-semibold uppercase tracking-wider bg-zinc-905 px-2 py-0.5 rounded border border-zinc-850">
                              {validationDetails.tabName}
                            </span>
                          </div>
                        )}

                        <div className="flex flex-col text-left gap-0.5">
                          <span className="text-[10px] text-zinc-500 uppercase tracking-wider">Original Gate Scan</span>
                          <span className="text-[11px] text-amber-500 font-medium">
                            {validationDetails.time}
                          </span>
                        </div>
                      </div>

                      {/* Duplicate Scan Lockout status bar */}
                      {cooldownRemaining > 0 && (
                        <div className="w-full mt-4 bg-zinc-950 border border-zinc-900 p-2.5 rounded-xl space-y-1">
                          <div className="flex justify-between items-center text-[9px] font-mono uppercase tracking-wider text-red-400">
                            <span>Lockout active - Anti-repeat lock</span>
                            <span className="font-bold">{cooldownRemaining}s remaining</span>
                          </div>
                          <div className="w-full bg-zinc-900 h-1.5 rounded-full overflow-hidden">
                            <div 
                              className="bg-red-500 h-full rounded-full transition-all duration-300 ease-out shadow-[0_0_8px_rgba(239,68,68,0.5)]"
                              style={{ width: `${(cooldownRemaining / cooldownSeconds) * 100}%` }}
                            ></div>
                          </div>
                        </div>
                      )}

                    </div>

                    <button 
                      onClick={() => {
                        setValidationState("idle"); 
                        startCameraScan();
                      }}
                      className="w-full bg-red-950/70 hover:bg-red-900/60 border border-red-800 text-stone-200 font-mono text-xs tracking-widest uppercase py-3 rounded-xl font-bold cursor-pointer transition-all shrink-0 flex justify-center items-center gap-2"
                    >
                      {cooldownRemaining > 0 ? (
                        <>
                          <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                          <span>AUTO-RESUMING IN {cooldownRemaining}S</span>
                        </>
                      ) : (
                        "TRY ORIGINAL CODE AGAIN"
                      )}
                    </button>
                  </div>
                )}

                 {/* C. NOT FOUND / INVALID: RED CARD */}
                 {validationState === "not_found" && (
                  <div className="w-full bg-black z-30 flex flex-col p-5 border-2 border-red-500 rounded-2xl relative">
                    <div className="flex-1 flex flex-col justify-center items-center text-center">
                      
                      <div className="w-14 h-14 bg-red-950/60 border border-red-500/40 rounded-full flex items-center justify-center mb-3 animate-bounce">
                        <AlertTriangle className="w-9 h-9 text-red-500" />
                      </div>
                      
                      <span className="font-mono text-xs bg-red-950/50 text-red-400 border border-red-500/20 px-3 py-1 rounded-full font-semibold uppercase tracking-widest leading-none mb-1">
                        NOT REGISTERED
                      </span>

                      <h2 className="font-display text-xl font-bold tracking-tight text-white mt-1 uppercase">
                        ❌ INVALID TICKET
                      </h2>

                      <p className="text-xs text-zinc-400 font-mono text-center mt-3 max-w-[240px]">
                        The scanned ticket code <span className="text-zinc-200 underline tracking-wider">{validationDetails.ticket}</span> under event <span className="text-red-400 font-bold tracking-widest uppercase">{validationDetails.eventCode}</span> is not found or has mismatched details.
                      </p>

                      {validationDetails.reason && (
                        <div className="mt-3 text-[11px] font-mono text-center text-red-400 bg-red-950/60 border border-red-500/20 px-3 py-2 rounded-lg max-w-[260px] leading-normal break-words">
                          <span className="text-red-300 font-bold uppercase block text-[9px] tracking-wider mb-0.5">Server Feedback</span>
                          {validationDetails.reason}
                        </div>
                      )}

                    </div>

                    <button 
                      onClick={() => {
                        setValidationState("idle"); 
                        startCameraScan();
                      }}
                      className="w-full bg-red-600 hover:bg-emerald-500 text-white font-mono text-xs tracking-widest uppercase py-3 rounded-xl font-bold cursor-pointer transition-all shrink-0"
                    >
                      RETRY CAMERA SCANNER
                    </button>
                  </div>
                )}

                {/* D. ERROR OVERLAY */}
                {validationState === "error" && (
                  <div className="w-full bg-zinc-950 z-30 flex flex-col p-5 border border-zinc-800 rounded-2xl relative">
                    <div className="flex-1 flex flex-col justify-center items-center text-center">
                      <div className="w-12 h-12 bg-red-950/40 border border-red-900/40 rounded-full flex items-center justify-center mb-3">
                        <AlertTriangle className="w-6 h-6 text-red-400" />
                      </div>
                      <h3 className="text-white text-sm font-bold uppercase tracking-wider font-mono">
                        NETWORK CONFIGURATION EXCEPTION
                      </h3>
                      <p className="text-zinc-500 text-xs mt-3 bg-black border border-zinc-900 p-3 rounded-lg text-left font-mono break-words leading-relaxed max-w-sm w-full">
                        {validationDetails.errorMessage || "Check Settings config."}
                      </p>
                    </div>

                    <div className="grid grid-cols-2 gap-2 shrink-0">
                      <button 
                        onClick={() => {
                          setValidationState("idle");
                          setActiveTab("settings");
                        }}
                        className="bg-zinc-900 hover:bg-zinc-800 text-stone-200 font-mono text-[10px] tracking-widest uppercase py-3 rounded-xl font-semibold border border-zinc-800"
                      >
                        FIX SETTINGS
                      </button>
                      <button 
                        onClick={() => {
                          if (scannedResultCode) {
                            processTicketVerification(scannedResultCode);
                          } else {
                            setValidationState("idle");
                          }
                        }}
                        className="bg-white hover:bg-stone-200 text-black font-mono text-[10px] tracking-widest uppercase py-3 rounded-xl font-bold"
                      >
                        RETRY RE-SYNC
                      </button>
                    </div>
                  </div>
                )}

              </div>

              {/* CAMERA INTERACTION BAR */}
              <div className="space-y-3">
                
                {/* BUTTONS FOR SCANNING AND UPLOADING */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  {!isScanning ? (
                    <button
                      onClick={startCameraScan}
                      className="w-full bg-white hover:bg-stone-100 text-black py-3.5 px-4 rounded-xl flex items-center justify-center gap-2.5 font-mono text-xs font-bold tracking-widest uppercase transition-transform active:scale-[0.99] cursor-pointer"
                    >
                      <Camera className="w-4 h-4 text-black" />
                      Activate Camera
                    </button>
                  ) : (
                    <button
                      onClick={stopCameraScan}
                      className="w-full bg-zinc-900 hover:bg-zinc-800 border border-zinc-800 text-red-500 py-3.5 px-4 rounded-xl flex items-center justify-center gap-2.5 font-mono text-xs font-bold tracking-widest uppercase transition-all"
                    >
                      <CameraOff className="w-4 h-4 text-red-500 animate-pulse" />
                      Close Camera
                    </button>
                  )}

                  <div className="relative">
                    <input 
                      type="file"
                      id="qr-file-uploader"
                      accept="image/*"
                      onChange={handleFileChange}
                      className="hidden"
                    />
                    <button
                      onClick={() => document.getElementById("qr-file-uploader")?.click()}
                      className="w-full bg-zinc-950 hover:bg-zinc-900 border border-zinc-800 text-stone-200 py-3.5 px-4 rounded-xl flex items-center justify-center gap-2.5 font-mono text-xs font-bold tracking-widest uppercase transition-transform active:scale-[0.99] cursor-pointer"
                    >
                      <Upload className="w-4 h-4 text-zinc-400" />
                      Upload QR Image
                    </button>
                  </div>
                </div>

                {/* MANUAL CODES ENTRY ACCORDION IF CAMERA FAILS */}
                <div className="bg-zinc-950 border border-zinc-900 p-3 rounded-xl">
                  <span className="text-[10px] font-mono uppercase text-zinc-500 font-bold block mb-2">
                    Or Manual Gate Check-In (Enter Ticket)
                  </span>
                  
                  <div className="flex gap-2">
                    <input 
                      type="text"
                      placeholder="e.g. THOC-aB3d5Fg"
                      id="manual-ticket-input"
                      maxLength={20}
                      className="flex-1 bg-black border border-zinc-800 text-white text-xs font-mono rounded-lg px-3 py-2 placeholder:text-zinc-700 tracking-wider focus:outline-none focus:border-zinc-500 text-center"
                      onKeyDown={(e) => {
                        if (e.key === "Enter") {
                          const input = e.currentTarget.value.trim();
                          handleManualCheckIn(input);
                        }
                      }}
                    />
                    <button 
                      onClick={() => {
                        const inputVal = (document.getElementById("manual-ticket-input") as HTMLInputElement)?.value;
                        handleManualCheckIn(inputVal || "");
                      }}
                      className="bg-zinc-900 hover:bg-zinc-800 border border-zinc-800 text-white font-mono text-xs px-4 py-2 rounded-lg font-bold"
                    >
                      Verify
                    </button>
                  </div>
                </div>

                {/* SCANNER RESULT PANEL */}
                <div className="bg-zinc-950 border border-zinc-800 rounded-2xl p-4.5 space-y-3 font-mono text-xs text-left animate-fade-in mt-2">
                  <div className="text-zinc-500 uppercase tracking-widest text-[9px] border-b border-zinc-900 pb-2 text-center font-bold">
                    --------------------------------<br />
                    SCAN RESULT<br />
                    --------------------------------
                  </div>

                  {isDebugMode ? (
                    /* DEBUG ON: DETAILED RESULT DATA CARD */
                    <div className="space-y-3.5">
                      <div>
                        <span className="text-zinc-500 block text-[10px] font-bold">Raw QR:</span>
                        <span className="text-zinc-300 font-mono text-xs font-bold leading-relaxed block mt-0.5 select-all break-all bg-black px-2.5 py-1.5 rounded border border-zinc-900">
                          {scanResult ? scanResult.raw : "STRANGER-2026|THOC-XA6V32JK"}
                        </span>
                      </div>

                      <div>
                        <span className="text-zinc-500 block text-[10px] font-bold">Event Code:</span>
                        <span className="text-amber-500 font-mono text-xs font-bold mt-0.5 block bg-black px-2.5 py-1.5 rounded border border-zinc-900">
                          {scanResult ? scanResult.eventCode : "STRANGER-2026"}
                        </span>
                      </div>

                      <div>
                        <span className="text-zinc-500 block text-[10px] font-bold">Ticket Code:</span>
                        <span className="text-stone-100 font-mono text-xs font-bold mt-0.5 block bg-black px-2.5 py-1.5 rounded border border-zinc-900">
                          {scanResult ? scanResult.ticketCode : "THOC-XA6V32JK"}
                        </span>
                      </div>

                      <div>
                        <span className="text-zinc-500 block text-[10px] font-bold">Validation:</span>
                        <span className="text-stone-300 font-mono text-xs font-semibold mt-0.5 block bg-black px-2.5 py-1.5 rounded border border-zinc-900">
                          {scanResult ? scanResult.validationText : "QR Format Valid ✅"}
                        </span>
                      </div>

                      <div>
                        <span className="text-zinc-500 block text-[10px] font-bold">Status:</span>
                        <span className="text-zinc-300 font-mono text-xs font-medium mt-0.5 block bg-black px-2.5 py-1.5 rounded border border-zinc-900">
                          {!scanResult ? "Ready For Verification" : (
                            validationState === "success"
                              ? "Ticket Verified"
                              : validationState === "already_used"
                              ? "Already Used"
                              : validationState === "not_found"
                              ? "Invalid Ticket"
                              : validationState === "validating"
                              ? "Validating..."
                              : scanResult.statusText
                          )}
                        </span>
                      </div>
                    </div>
                  ) : (
                    /* DEBUG OFF: MINIMAL CARD */
                    <div className="py-5 text-center flex flex-col justify-center items-center">
                      {!scanResult ? (
                        <span className="text-zinc-600 text-xs tracking-wider">Awaiting scan or manual check-in...</span>
                      ) : (
                        <div className="space-y-1">
                          {validationState === "success" ? (
                            <span className="text-emerald-400 font-bold text-base tracking-widest block font-mono">
                              Ticket Verified
                            </span>
                          ) : (
                            <span className="text-red-500 font-bold text-base tracking-widest block font-mono">
                              Invalid Ticket
                            </span>
                          )}
                        </div>
                      )}
                    </div>
                  )}

                  <div className="text-zinc-500 uppercase tracking-widest text-[9px] pt-1 text-center font-bold">
                    --------------------------------
                  </div>
                </div>

              </div>

            </div>
          )}

          {/* ======================================= */}
          {/* TAB 2: QR GENERATOR SECTION               */}
          {/* ======================================= */}
          {activeTab === "generator" && (
            <div id="tab-generator-section" className="flex flex-col gap-4">
              
              <div className="flex justify-between items-center pb-2 border-b border-zinc-900">
                <div className="flex items-center gap-2">
                  <QrCode className="w-4 h-4 text-zinc-400" />
                  <h2 className="font-display text-sm uppercase tracking-wider font-semibold text-white">
                    Ticket Constructor
                  </h2>
                </div>
                <span className="text-[10px] bg-zinc-900 text-zinc-400 font-mono px-2 py-0.5 rounded-full border border-zinc-800">
                  Core QR Engine
                </span>
              </div>

              {/* GENERATOR CONTROLS */}
              <div className="space-y-4">
                
                <div>
                  <label className="text-[10px] font-mono tracking-widest text-zinc-500 uppercase block mb-1">
                    Event Code Scope Prefix <span className="text-red-500/80">*</span>
                  </label>
                  <input
                    type="text"
                    value={constructorEventCode}
                    onChange={(e) => setConstructorEventCode(e.target.value)}
                    placeholder="e.g. STRANGER-2026"
                    className="w-full bg-zinc-950 border border-zinc-900 rounded-lg px-3 py-2 text-xs font-mono tracking-wider text-amber-500 font-bold focus:outline-none focus:border-zinc-700"
                  />
                </div>

                <div>
                  <label className="text-[10px] font-mono tracking-widest text-zinc-500 uppercase block mb-1">
                    Guest Name <span className="text-zinc-700">(Optional Ticket Card Label)</span>
                  </label>
                  <div className="relative">
                    <input 
                      type="text"
                      placeholder="e.g. Liam Sterling"
                      value={ticketNameInput}
                      onChange={(e) => setTicketNameInput(e.target.value)}
                      className="w-full bg-zinc-950 border border-zinc-900 rounded-lg px-3 py-2 text-xs font-display text-stone-200 placeholder:text-zinc-700 tracking-wide focus:outline-none focus:border-zinc-700"
                      maxLength={40}
                    />
                    <User className="absolute right-3 top-2.5 w-3.5 h-3.5 text-zinc-700" />
                  </div>
                </div>

                <div>
                  <div className="flex justify-between items-center mb-1">
                    <label className="text-[10px] font-mono tracking-widest text-zinc-500 uppercase block">
                      Spreadsheet Tab (Tier) <span className="text-red-500/80">*</span>
                    </label>
                    <button
                      type="button"
                      onClick={() => fetchSheetNamesFromSpreadsheet()}
                      disabled={isFetchingSheets}
                      className="text-[9px] font-mono text-amber-500 hover:text-amber-400 flex items-center gap-1 active:scale-95 disabled:opacity-50 transition-all cursor-pointer"
                      title="Fetch dynamic sheet names from active Spreadsheet"
                    >
                      <RefreshCw className={`w-2.5 h-2.5 ${isFetchingSheets ? 'animate-spin' : ''}`} />
                      Fetch Tab List
                    </button>
                  </div>
                  <div className="relative">
                    <input
                      type="text"
                      list="constructor-tabs-list"
                      placeholder="e.g. Stranger_People"
                      value={constructorTabName}
                      onChange={(e) => setConstructorTabName(e.target.value)}
                      className="w-full bg-zinc-950 border border-zinc-900 rounded-lg px-3 py-2 text-xs font-mono text-stone-200 tracking-wide focus:outline-none focus:border-zinc-700"
                    />
                    <datalist id="constructor-tabs-list">
                      {fetchedSheetTabs.map(tab => (
                        <option key={tab} value={tab} />
                      ))}
                    </datalist>
                  </div>
                  <div className="flex flex-wrap gap-1 mt-1.5">
                    <span className="text-[8px] text-zinc-600 self-center">Presets:</span>
                    {fetchedSheetTabs.map(preset => (
                      <button
                        key={preset}
                        type="button"
                        onClick={() => setConstructorTabName(preset)}
                        className={`text-[8px] font-mono px-1.5 py-0.5 rounded border transition-all ${
                          constructorTabName === preset
                            ? "bg-amber-950/40 text-amber-500 border-amber-900/40 font-bold"
                            : "bg-zinc-950 text-zinc-500 border-zinc-900 hover:text-zinc-350"
                        }`}
                      >
                        {preset}
                      </button>
                    ))}
                  </div>
                </div>

                <div>
                  <label className="text-[10px] font-mono tracking-widest text-zinc-500 uppercase block mb-1">
                    Strict Ticket Number <span className="text-red-500/80">*</span>
                  </label>
                  
                  <div className="flex flex-col gap-2">
                    <div className="relative w-full">
                      <input 
                        type="text"
                        placeholder="e.g. THOC-aB3d5Fg"
                        value={inputTicket}
                        onChange={(e) => setInputTicket(e.target.value)}
                        onBlur={() => {
                          const val = inputTicket.trim();
                          const prefix = constructorEventCode ? constructorEventCode.split('-')[0] + "-" : "THOC-";
                          if (val && val !== prefix) {
                            handlePullNameFromSpreadsheet(val, true);
                          }
                        }}
                        onKeyDown={(e) => {
                          if (e.key === "Enter") {
                            e.preventDefault();
                            handlePullNameFromSpreadsheet(inputTicket, false);
                          }
                        }}
                        className="w-full bg-zinc-950 border border-zinc-900 rounded-lg px-4 py-3 text-lg sm:text-xl font-mono text-white tracking-widest placeholder:text-zinc-700 focus:outline-none focus:border-zinc-700 text-center"
                        maxLength={20}
                      />
                    </div>
                    
                    <div className="flex gap-2">
                      <button
                        onClick={() => handlePullNameFromSpreadsheet(inputTicket, false)}
                        disabled={isSyncingName}
                        className="flex-1 bg-amber-950/30 hover:bg-amber-900/40 text-amber-500 font-mono text-xs px-2.5 py-2 rounded-lg border border-amber-900/40 flex items-center justify-center gap-1.5 font-medium active:scale-95 disabled:opacity-50 transition-all cursor-pointer"
                        title="Look up and pull the Guest Name from the spreadsheet roster matching this Ticket number."
                      >
                        {isSyncingName ? (
                          <RefreshCw className="w-3 h-3 animate-spin" />
                        ) : (
                          <Database className="w-3 h-3" />
                        )}
                        Fetch Name
                      </button>

                      <button
                        onClick={generateRandomTicketCode}
                        className="flex-1 bg-zinc-900 hover:bg-zinc-800 text-stone-300 font-mono text-xs px-2.5 py-2 rounded-lg border border-zinc-800 flex items-center justify-center gap-1 font-medium active:scale-95 transition-all cursor-pointer"
                        title="Generate a randomized ticket number containing capital, small, and digital characters."
                      >
                        <RefreshCw className="w-3 h-3 text-zinc-400" />
                        Auto Rand
                      </button>
                    </div>
                  </div>
                  <span className="text-[9px] font-mono text-zinc-600 block mt-2 tracking-wider leading-relaxed">
                    Must be 7 to 20 characters containing any combination of capital, small, and digital characters (e.g. mixed-case and numbers).
                  </span>
                </div>

                <button
                  onClick={handleGenerateQR}
                  disabled={isGenerating}
                  className="w-full bg-white hover:bg-zinc-200 text-black py-3 rounded-lg font-mono text-xs font-bold tracking-widest uppercase transition-all"
                >
                  {isGenerating ? "Assembling Core Vectors..." : "Assemble & Generate QR"}
                </button>

              </div>

              {/* RENDER STAGE (Luxurious physical Boarding Pass representation) */}
              {generatedQrUrl && (
                <div className="bg-gradient-to-b from-zinc-950 to-black rounded-2xl border border-zinc-800 p-4 shadow-xl flex flex-col items-center">
                  
                  {/* Decorative luxury boarding pass design head */}
                  <div className="w-full border-b border-dashed border-zinc-800 pb-3 flex flex-col items-center">
                    <span className="font-serif text-[11px] font-bold tracking-[0.2em] text-white">
                      THE HOUSE OF CONNECTIONS
                    </span>
                    <span className="text-[8px] font-display uppercase tracking-[0.3em] text-zinc-500 mt-0.5">
                      OFFICIAL ADMISSION TICKET
                    </span>
                  </div>

                  <div className="bg-white p-3.5 my-4 rounded-xl shadow-lg">
                    {/* Render Image Source */}
                    <img 
                      src={generatedQrUrl} 
                      alt="Pass Ticket QR Code Representation" 
                      className="w-44 h-44 object-contain rounded"
                    />
                  </div>

                  {/* Pass details */}
                  <div className="w-full space-y-2 font-mono text-[10px] text-zinc-400 uppercase tracking-wider mb-4 border-t border-dashed border-zinc-800 pt-3">
                    <div className="flex justify-between">
                      <span className="text-zinc-500">Host Event:</span>
                      <span className="text-stone-300 font-bold max-w-[180px] truncate text-right lowercase first-letter:uppercase">
                        {eventName}
                      </span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-zinc-500">Event Code:</span>
                      <span className="text-amber-500 font-bold text-right tracking-widest">
                        {constructorEventCode}
                      </span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-zinc-500">Sheet Tab:</span>
                      <span className="text-stone-300 font-mono font-bold tracking-wider text-right uppercase">
                        {constructorTabName}
                      </span>
                    </div>
                    {ticketNameInput && (
                      <div className="flex justify-between">
                        <span className="text-zinc-500">Holder:</span>
                        <span className="text-zinc-300 font-bold italic normal-case">{ticketNameInput}</span>
                      </div>
                    )}
                    <div className="flex justify-between">
                      <span className="text-zinc-500">Ticket ID:</span>
                      <span className="text-white font-bold tracking-widest">
                        {inputTicket}
                      </span>
                    </div>
                    <div className="flex flex-col text-left gap-1 border-t border-zinc-900 pt-2 text-[9px]">
                      <span className="text-zinc-600 block uppercase font-bold tracking-wider">Raw QR Data String:</span>
                      <span className="text-amber-500 font-bold break-all font-mono tracking-widest">
                        {constructorEventCode}|{constructorTabName}|{inputTicket}
                      </span>
                    </div>
                  </div>

                  {/* Actions */}
                  <button
                    onClick={downloadQRAsPNG}
                    className="w-full bg-zinc-900 hover:bg-zinc-800 text-white font-mono text-xs tracking-widest uppercase py-3 rounded-xl font-bold border border-zinc-800 flex items-center justify-center gap-2"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Download PNG Ticket
                  </button>

                </div>
              )}

            </div>
          )}

          {/* ======================================= */}
          {/* TAB 3: SETTINGS SECTION                  */}
          {/* ======================================= */}
          {activeTab === "settings" && (
            <div id="tab-settings-section" className="flex flex-col gap-4">
              
              <div className="flex justify-between items-center pb-2 border-b border-zinc-900">
                <div className="flex items-center gap-2">
                  <Settings className="w-4 h-4 text-zinc-400" />
                  <h2 className="font-display text-sm uppercase tracking-wider font-semibold text-white">
                    System Configuration
                  </h2>
                </div>
              </div>

              {/* DEBUG MODE SWITCH */}
              <div className="bg-zinc-950 border border-zinc-900 rounded-xl p-3.5 space-y-3">
                <div className="flex justify-between items-center">
                  <div>
                    <h3 className="text-xs font-mono font-bold tracking-wider uppercase text-white">
                      Enable Debug Mode
                    </h3>
                    <p className="text-[10px] text-zinc-500 mt-0.5 max-w-[260px]">
                      Show Raw QR Data, Event Code, Ticket Code, and Validation result details on verification.
                    </p>
                  </div>
                  <label className="relative inline-flex items-center cursor-pointer">
                    <input 
                      type="checkbox" 
                      className="sr-only peer" 
                      id="thoc-debug-switch"
                      checked={isDebugMode}
                      onChange={(e) => {
                        setIsDebugMode(e.target.checked);
                        localStorage.setItem("thoc_debug_mode", String(e.target.checked));
                        showToastNotification(`Debug Mode ${e.target.checked ? "Enabled" : "Disabled"}`);
                      }}
                    />
                    <div className="w-9 h-5 bg-zinc-800 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-zinc-400 after:border-zinc-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-white peer-checked:after:bg-black"></div>
                  </label>
                </div>
              </div>

              {/* SHEET SETTINGS FIELDS */}
              <div className="space-y-3 bg-zinc-950 border border-zinc-900 rounded-xl p-3.5">
                <h3 className="text-xs font-mono font-bold tracking-wider uppercase text-white border-b border-zinc-900 pb-2">
                  Live Event Integration Keys
                </h3>

                <div>
                  <label className="text-[10px] font-mono uppercase text-zinc-500 block mb-1">
                    Event Display Title
                  </label>
                  <input
                    type="text"
                    id="conf-event-name"
                    placeholder="e.g. THoC Summit"
                    defaultValue={eventName}
                    className="w-full bg-black border border-zinc-800 text-stone-200 text-xs px-3 py-2 rounded-lg focus:outline-none focus:border-zinc-500 font-display"
                  />
                </div>

                <div>
                  <label className="text-[10px] font-mono uppercase text-zinc-500 block mb-1">
                    Event Code Scope Prefix
                  </label>
                  <input
                    type="text"
                    id="conf-event-code"
                    placeholder="e.g. STRANGER-2026"
                    defaultValue={eventCode}
                    className="w-full bg-black border border-zinc-800 text-stone-200 text-xs px-3 py-2 rounded-lg focus:outline-none focus:border-zinc-500 font-mono tracking-wider text-amber-500 font-bold"
                  />
                  <span className="text-[8px] text-zinc-600 block mt-1 tracking-wider leading-relaxed">
                    Used to prefix generated and verified barcodes like: <code><b>[EVENT_CODE]</b>|THOC-IJS3S7BU</code>
                  </span>
                </div>

                <div>
                  <label className="text-[10px] font-mono uppercase text-zinc-500 block mb-1">
                    Google Spreadsheet Sheet ID
                  </label>
                  <input
                    type="text"
                    id="conf-sheet-id"
                    placeholder="1_xxxxx_xxxxxxxxxxxxxxxxxxxxxxxxxx"
                    defaultValue={sheetId}
                    className="w-full bg-black border border-zinc-800 text-stone-200 text-xs px-3 py-2 rounded-lg focus:outline-none focus:border-zinc-500 font-mono"
                  />
                  <span className="text-[8px] text-zinc-600 block mt-1 tracking-wider leading-relaxed">
                    The long code inside your Google Sheet URL: <code>docs.google.com/spreadsheets/d/<b>[SHEET_ID]</b>/edit</code>
                  </span>
                </div>

                <div>
                  <label className="text-[10px] font-mono uppercase text-zinc-500 block mb-1">
                    Default Fallback Sheet Tab
                  </label>
                  <input
                    type="text"
                    id="conf-fallback-tab-name"
                    list="conf-fallback-tabs-datalist"
                    placeholder="e.g. Stranger_People"
                    defaultValue={fallbackTabName}
                    onChange={(e) => setFallbackTabName(e.target.value)}
                    className="w-full bg-black border border-zinc-800 text-stone-200 text-xs px-3 py-2 rounded-lg focus:outline-none focus:border-zinc-500 font-mono"
                  />
                  <datalist id="conf-fallback-tabs-datalist">
                    {fetchedSheetTabs.map(tab => (
                      <option key={tab} value={tab} />
                    ))}
                  </datalist>
                  <span className="text-[8px] text-zinc-600 block mt-1 tracking-wider leading-relaxed">
                    The default tab to use when the scanned QR code does not contain a tab name middle segment (e.g. legacy/2-part QR codes). Examples: <code>Stranger_People</code>, <code>Stranger_Creators</code>, <code>Pickleball_People</code>, etc.
                  </span>
                </div>

                {/* DYNAMIC SHETS/TABS FETCH PANEL */}
                <div className="bg-zinc-950 border border-zinc-900 rounded-xl p-3.5 mt-2 space-y-2.5">
                  <div className="flex justify-between items-center border-b border-zinc-900 pb-1.5">
                    <h4 className="text-[11px] font-mono font-bold tracking-wider uppercase text-white">
                      Live Spreadsheet Tabs
                    </h4>
                    <button
                      type="button"
                      onClick={() => fetchSheetNamesFromSpreadsheet()}
                      disabled={isFetchingSheets}
                      className="bg-zinc-900 hover:bg-zinc-800 border border-zinc-800 hover:border-zinc-700 text-amber-500 font-mono text-[9px] px-2 py-0.5 rounded flex items-center gap-1 active:scale-95 transition-all cursor-pointer disabled:opacity-50"
                    >
                      <RefreshCw className={`w-2.5 h-2.5 ${isFetchingSheets ? 'animate-spin' : ''}`} />
                      Fetch Tab Names
                    </button>
                  </div>
                  
                  <span className="text-[9px] text-zinc-500 block leading-normal font-mono">
                    Dynamically pull tab/sheet names from your active master spreadsheet to synchronize with dynamic attendee groups. Custom write any sheets into form fields above.
                  </span>

                  <div className="flex flex-wrap gap-1 mt-1">
                    {fetchedSheetTabs.map((tab) => (
                      <button
                        key={tab}
                        type="button"
                        onClick={() => {
                          const elInput = document.getElementById("conf-fallback-tab-name") as HTMLInputElement;
                          if (elInput) {
                            elInput.value = tab;
                          }
                          setFallbackTabName(tab);
                          localStorage.setItem("thoc_fallback_tab_name", tab);
                          showToastNotification(`Set default fallback tab to "${tab}". Remember to Save Configuration below!`);
                        }}
                        className={`text-[8px] font-mono px-2 py-0.5 rounded border transition-all ${
                          fallbackTabName === tab
                            ? "bg-amber-955 text-amber-500 border-amber-900/40 font-bold"
                            : "bg-zinc-950 text-zinc-500 border-zinc-900 hover:text-zinc-350"
                        }`}
                      >
                        {tab} {fallbackTabName === tab ? "★" : ""}
                      </button>
                    ))}
                  </div>
                </div>

                <div>
                  <label className="text-[10px] font-mono uppercase text-zinc-500 block mb-1">
                    Google Apps Script API Endpoint URL
                  </label>
                  <input
                    type="url"
                    id="conf-script-url"
                    placeholder="https://script.google.com/macros/s/..../exec"
                    defaultValue={scriptUrl}
                    className="w-full bg-black border border-zinc-800 text-stone-200 text-xs px-3 py-2 rounded-lg focus:outline-none focus:border-zinc-500 font-mono"
                  />
                  <span className="text-[8px] text-zinc-600 block mt-1 tracking-wider leading-relaxed">
                    Obtained after clicking <b>Deploy &gt; New deployment &gt; Web app</b> inside Apps Script Editor. Configure "Execute as: Me" and "Who has access: Anyone".
                  </span>
                </div>

                <div>
                  <label className="text-[10px] font-mono uppercase text-zinc-500 block mb-1">
                    Scan Cooldown Period
                  </label>
                  <select
                    id="conf-cooldown-seconds"
                    defaultValue={cooldownSeconds}
                    className="w-full bg-black border border-zinc-800 text-stone-200 text-xs px-2.5 py-2 rounded-lg focus:outline-none focus:border-zinc-500 font-mono uppercase tracking-wider text-amber-500 font-bold text-[11px]"
                  >
                    <option value={3}>3 Seconds Cooldown</option>
                    <option value={4}>4 Seconds Cooldown</option>
                    <option value={5}>5 Seconds Cooldown</option>
                    <option value={6}>6 Seconds Cooldown</option>
                    <option value={8}>8 Seconds Cooldown</option>
                    <option value={10}>10 Seconds Cooldown</option>
                  </select>
                  <span className="text-[8px] text-zinc-600 block mt-1 tracking-wider leading-relaxed">
                    Pause duration after a successful lock scan before the lens resumes. Repeats of the same ticket will be strictly blocked during this threshold.
                  </span>
                </div>

                {/* SAVE BUTTON */}
                <button
                  onClick={() => {
                    const elName = (document.getElementById("conf-event-name") as HTMLInputElement)?.value;
                    const elCode = (document.getElementById("conf-event-code") as HTMLInputElement)?.value;
                    const elSheet = (document.getElementById("conf-sheet-id") as HTMLInputElement)?.value;
                    const elUrl = (document.getElementById("conf-script-url") as HTMLInputElement)?.value;
                    const elFallbackTab = (document.getElementById("conf-fallback-tab-name") as HTMLInputElement)?.value;
                    const elCooldown = parseInt((document.getElementById("conf-cooldown-seconds") as HTMLSelectElement)?.value || "4", 10);
                    
                    saveConfiguration(
                      elName || "THE HOUSE OF CONNECTIONS Event",
                      elCode || "STRANGER-2026",
                      elSheet || "",
                      elUrl || "",
                      elFallbackTab || "Stranger_People",
                      elCooldown
                    );
                  }}
                  className="w-full bg-white hover:bg-stone-200 text-black font-mono text-xs tracking-widest uppercase font-bold text-center py-2.5 rounded-lg border border-transparent mt-2 cursor-pointer transition-all"
                >
                  Save Configuration Sync
                </button>
              </div>

              {/* EXCEL/GOOGLE SHEETS SCRIPT COMPLIANCE INSTRUCTION BOX */}
              <div className="bg-zinc-950 border border-zinc-900 rounded-xl p-3.5">
                <button 
                  onClick={() => setShowConfigCode(!showConfigCode)}
                  className="w-full flex justify-between items-center font-mono text-xs uppercase tracking-wider text-stone-300"
                >
                  <div className="flex items-center gap-2">
                    <Database className="w-3.5 h-3.5 text-zinc-400" />
                    <span>How to Setup Google Sheet (API Script)</span>
                  </div>
                  <span className="text-[10px] text-zinc-500 font-bold">
                    {showConfigCode ? "[ HIDE CODE ]" : "[ SHOW SETUP ]"}
                  </span>
                </button>

                {showConfigCode && (
                  <div className="mt-3 text-[10px] font-mono text-zinc-400 leading-relaxed space-y-3.5 border-t border-zinc-900 pt-3">
                    <p>
                      1. Open your <b>Google Sheet</b>. Create headers in Row 1: <b>"Ticket No"</b>, <b>"Status"</b>, <b>"Event Code"</b>, and <b>"Check In Time"</b>. You can also optionally add <b>"Event"</b> and <b>"Name"</b> headers.
                    </p>
                    <p>
                      2. Add at least one test row with a real ticket number, matching event code, event title, guest name, status (e.g., leave empty or add "Available"), and check in time.
                    </p>
                    <p>
                      3. Click <b>Extensions &gt; Apps Script</b> in GSheets. Erase all default code, paste the script below, save, and click <b>Deploy &gt; New deployment &gt; Web app</b> (execute as "Me", access is "Anyone").
                    </p>

                    <div className="bg-stone-950 p-2.5 rounded-lg border border-zinc-900 relative">
                      <pre className="overflow-x-auto text-[8px] leading-tight text-white max-h-52">{GOOGLE_APPS_SCRIPT_CODE}</pre>
                      <button 
                        onClick={() => {
                          navigator.clipboard.writeText(GOOGLE_APPS_SCRIPT_CODE);
                          showToastNotification("Apps Script Code Copied");
                        }}
                        className="absolute right-2 top-2 bg-zinc-900 hover:bg-zinc-800 border border-zinc-800 p-1.5 rounded text-white cursor-pointer transition-all active:scale-95"
                        title="Copy complete Apps Script code to clipboard"
                      >
                        <Copy className="w-3 h-3" />
                      </button>
                    </div>
                  </div>
                )}
              </div>

            </div>
          )}

        </div>

        {/* ======================================= */}
        {/* PHYSICAL DASHBOARD NAV BAR AT THE BOTTOM */}
        {/* ======================================= */}
        <nav id="thoc-bottom-selector" className="bg-[#0b0b0b] border-t border-zinc-900 flex justify-around items-stretch py-2 h-16 shrink-0 z-10">
          
          {/* TAB 1 ACTION */}
          <button
            onClick={() => {
              setActiveTab("scanner");
              setValidationState("idle");
            }}
            className={`flex-1 flex flex-col items-center justify-center gap-1 cursor-pointer transition-all ${
              activeTab === "scanner" ? "text-white font-bold" : "text-zinc-600 hover:text-zinc-400"
            }`}
          >
            <Scan className={`w-4 h-4 transition-transform ${activeTab === "scanner" ? "scale-110 text-white" : ""}`} />
            <span className="font-display text-[9px] uppercase tracking-wider">Gate Scanner</span>
          </button>

          {/* TAB 2 ACTION */}
          <button
            onClick={() => {
              setActiveTab("generator");
            }}
            className={`flex-1 flex flex-col items-center justify-center gap-1 cursor-pointer transition-all ${
              activeTab === "generator" ? "text-white font-bold" : "text-zinc-600 hover:text-zinc-400"
            }`}
          >
            <QrCode className={`w-4 h-4 transition-transform ${activeTab === "generator" ? "scale-110 text-white" : ""}`} />
            <span className="font-display text-[9px] uppercase tracking-wider">Constructor</span>
          </button>

          {/* TAB 3 ACTION */}
          <button
            onClick={() => {
              setActiveTab("settings");
            }}
            className={`flex-1 flex flex-col items-center justify-center gap-1 cursor-pointer transition-all ${
              activeTab === "settings" ? "text-white font-bold" : "text-zinc-600 hover:text-zinc-400"
            }`}
          >
            <Settings className={`w-4 h-4 transition-transform ${activeTab === "settings" ? "scale-110 text-white" : ""}`} />
            <span className="font-display text-[9px] uppercase tracking-wider">Settings</span>
          </button>

        </nav>

      </main>

      {/* FLOATING SYSTEM TOAST SYSTEM */}
      {toastMessage && (
        <div className="fixed bottom-6 z-50 left-1/2 -translate-x-1/2 bg-zinc-950 border border-zinc-800 px-4 py-2 rounded-xl flex items-center gap-2 shadow-lg max-w-[280px] w-auto">
          <div className="w-1.5 h-1.5 rounded-full bg-white animate-ping"></div>
          <span className="font-mono text-[9px] uppercase tracking-wider font-semibold text-stone-100">
            {toastMessage}
          </span>
        </div>
      )}

      {/* FOOTER DESCRIPTOR */}
      <footer className="w-full max-w-md text-center py-4 mt-1">
        <p className="font-mono text-[8px] uppercase tracking-[0.3em] text-zinc-700">
          DESIGNED FOR THE HOUSE OF CONNECTIONS BY DEEPMIND BUILDERS · SECURE OFFLINE CORE
        </p>
      </footer>

    </div>
  );
}
