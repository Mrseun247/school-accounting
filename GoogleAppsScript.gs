// Google Apps Script for School Accounting System
// This script uses Google Sheets as a database
// Deploy as a web app: Deploy > New Deployment > Web App

// Sheet names - Update these to match your Google Sheet tabs
const STUDENTS_SHEET = "Students";
const FEES_SHEET = "Fees";
const PAYMENTS_SHEET = "Payments";
const FEE_BREAKDOWN_SHEET = "Fee Breakdown";

// Initialize sheets
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create Students sheet if it doesn't exist
  if (!ss.getSheetByName(STUDENTS_SHEET)) {
    const studentsSheet = ss.insertSheet(STUDENTS_SHEET);
    studentsSheet.appendRow(["Student ID", "Name", "Class", "Gender", "Parent/Guardian Name", "Date Created"]);
  }
  
  // Create Fees sheet if it doesn't exist
  if (!ss.getSheetByName(FEES_SHEET)) {
    const feesSheet = ss.insertSheet(FEES_SHEET);
    feesSheet.appendRow(["Fee ID", "Term", "Class", "Total Amount", "Date Created"]);
  }
  
  // Create Fee Breakdown sheet if it doesn't exist
  if (!ss.getSheetByName(FEE_BREAKDOWN_SHEET)) {
    const breakdownSheet = ss.insertSheet(FEE_BREAKDOWN_SHEET);
    breakdownSheet.appendRow(["Fee ID", "Class", "Tuition", "Uniform", "Exam", "Club Activity", "Books", "PTA", "ICT", "Coding", "Library", "Lab"]);
  }
  
  // Create Payments sheet if it doesn't exist
  if (!ss.getSheetByName(PAYMENTS_SHEET)) {
    const paymentsSheet = ss.insertSheet(PAYMENTS_SHEET);
    paymentsSheet.appendRow(["Payment ID", "Student ID", "Fee ID", "Amount Paid", "Payment Date", "Date Recorded"]);
  }
}

// Main doGet function - handles all requests
function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback; // JSONP callback
  
  let result;
  
  try {
    switch(action) {
      case 'addStudent':
        result = addStudent(e.parameter);
        break;
      case 'addFee':
        result = addFee(e.parameter);
        break;
      case 'addPayment':
        result = addPayment(e.parameter);
        break;
      case 'getStudentFees':
        result = getStudentFees(e.parameter.studentID);
        break;
      case 'searchStudent':
        result = searchStudent(e.parameter.query);
        break;
      case 'getReport':
        result = getReport(e.parameter.term);
        break;
      default:
        result = { success: false, message: "Unknown action" };
    }
  } catch(error) {
    result = { success: false, message: error.toString() };
  }
  
  // Return JSONP response
  return ContentService.createTextOutput(
    callback + "(" + JSON.stringify(result) + ")"
  ).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// Add a new student
function addStudent(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STUDENTS_SHEET);
  
  if (!sheet) {
    return { success: false, message: "Students sheet not found. Run initializeSheets() first." };
  }
  
  // Validate inputs
  if (!params.name || !params.class || !params.gender || !params.parentName) {
    return { success: false, message: "Missing required fields: name, class, gender, parentName" };
  }
  
  // Generate custom Student ID (VMCS001, VMCS002, etc.)
  let maxNumber = 0;
  const prefix = "VMCS";
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const id = data[i][0].toString();
    if (id.startsWith(prefix)) {
      const numberPart = parseInt(id.substring(prefix.length));
      if (numberPart > maxNumber) {
        maxNumber = numberPart;
      }
    }
  }
  
  const nextNumber = maxNumber + 1;
  const studentID = prefix + String(nextNumber).padStart(3, '0'); // VMCS001, VMCS002, etc.
  
  // Add new student
  sheet.appendRow([
    studentID,
    params.name,
    params.class,
    params.gender,
    params.parentName,
    new Date()
  ]);
  
  return { 
    success: true, 
    message: "Student registered successfully!",
    studentID: studentID
  };
}

// Add a new fee
function addFee(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FEES_SHEET);
  const breakdownSheet = ss.getSheetByName(FEE_BREAKDOWN_SHEET);
  
  if (!sheet || !breakdownSheet) {
    return { success: false, message: "Fees sheets not found. Run initializeSheets() first." };
  }
  
  // Validate inputs
  if (!params.term || !params.class || !params.amount) {
    return { success: false, message: "Missing required fields: term, class, amount" };
  }
  
  // Check if fee already exists for this term and class
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === params.term && data[i][2] === params.class) {
      return { success: false, message: "Fee for " + params.class + " in " + params.term + " already exists" };
    }
  }
  
  // Generate Fee ID
  const feeID = "FEE" + Date.now();
  
  // Add new fee to Fees sheet
  sheet.appendRow([
    feeID,
    params.term,
    params.class,
    parseFloat(params.amount),
    new Date()
  ]);
  
  // Add fee breakdown if provided
  if (params.breakdown) {
    try {
      const breakdown = JSON.parse(params.breakdown);
      breakdownSheet.appendRow([
        feeID,
        params.class,
        breakdown.tuition || 0,
        breakdown.uniform || 0,
        breakdown.exam || 0,
        breakdown.club || 0,
        breakdown.books || 0,
        breakdown.pta || 0,
        breakdown.ict || 0,
        breakdown.coding || 0,
        breakdown.library || 0,
        breakdown.lab || 0
      ]);
    } catch(e) {
      // If breakdown parsing fails, just continue without it
    }
  }
  
  return { 
    success: true, 
    message: "Fee registered successfully for " + params.class + "!",
    feeID: feeID
  };
}

// Record a payment
function addPayment(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentsSheet = ss.getSheetByName(PAYMENTS_SHEET);
  const studentsSheet = ss.getSheetByName(STUDENTS_SHEET);
  
  if (!paymentsSheet || !studentsSheet) {
    return { success: false, message: "Required sheets not found. Run initializeSheets() first." };
  }
  
  // Validate inputs
  if (!params.studentID || !params.amountPaid || !params.date || !params.feeItems) {
    return { success: false, message: "Missing required fields: studentID, feeItems, amountPaid, date" };
  }
  
  // Verify student exists
  const studentData = studentsSheet.getDataRange().getValues();
  let studentFound = false;
  for (let i = 1; i < studentData.length; i++) {
    if (studentData[i][0] === params.studentID) {
      studentFound = true;
      break;
    }
  }
  
  if (!studentFound) {
    return { success: false, message: "Student not found" };
  }
  
  // Generate Payment ID
  const paymentID = "PAY" + Date.now();
  
  // Add new payment
  paymentsSheet.appendRow([
    paymentID,
    params.studentID,
    params.feeItems,  // Store which fee items were paid
    parseFloat(params.amountPaid),
    params.date,
    new Date()
  ]);
  
  return { 
    success: true, 
    message: "Payment recorded successfully!",
    paymentID: paymentID
  };
}

// Get student fees
function getStudentFees(studentID) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const studentsSheet = ss.getSheetByName(STUDENTS_SHEET);
    const feesSheet = ss.getSheetByName(FEES_SHEET);
    const breakdownSheet = ss.getSheetByName(FEE_BREAKDOWN_SHEET);
    
    if (!studentsSheet) {
      return { success: false, message: "Students sheet not found" };
    }
    if (!feesSheet) {
      return { success: false, message: "Fees sheet not found" };
    }
    if (!breakdownSheet) {
      return { success: false, message: "Fee Breakdown sheet not found" };
    }
    
    // Find student
    const studentData = studentsSheet.getDataRange().getValues();
    let student = null;
    
    for (let i = 1; i < studentData.length; i++) {
      if (studentData[i][0] == studentID) {
        student = {
          id: studentData[i][0],
          name: studentData[i][1],
          class: studentData[i][2],
          gender: studentData[i][3],
          parentName: studentData[i][4]
        };
        break;
      }
    }
    
    if (!student) {
      return { success: false, message: "Student not found with ID: " + studentID };
    }
    
    // Get fees for this student's class
    const feesData = feesSheet.getDataRange().getValues();
    let feeID = null;
    
    // Find the latest fee for this student's class
    for (let i = feesData.length - 1; i >= 1; i--) {
      if (feesData[i][2] == student.class) {
        feeID = feesData[i][0];
        break;
      }
    }
    
    if (!feeID) {
      return { success: false, message: "No fees registered for class: " + student.class };
    }
    
    // Get fee breakdown
    const breakdownData = breakdownSheet.getDataRange().getValues();
    let fees = {
      tuition: 0,
      uniform: 0,
      exam: 0,
      club: 0,
      books: 0,
      pta: 0,
      ict: 0,
      coding: 0,
      library: 0,
      lab: 0
    };
    
    for (let i = 1; i < breakdownData.length; i++) {
      if (breakdownData[i][0] == feeID) {
        fees.tuition = parseFloat(breakdownData[i][2]) || 0;
        fees.uniform = parseFloat(breakdownData[i][3]) || 0;
        fees.exam = parseFloat(breakdownData[i][4]) || 0;
        fees.club = parseFloat(breakdownData[i][5]) || 0;
        fees.books = parseFloat(breakdownData[i][6]) || 0;
        fees.pta = parseFloat(breakdownData[i][7]) || 0;
        fees.ict = parseFloat(breakdownData[i][8]) || 0;
        fees.coding = parseFloat(breakdownData[i][9]) || 0;
        fees.library = parseFloat(breakdownData[i][10]) || 0;
        fees.lab = parseFloat(breakdownData[i][11]) || 0;
        break;
      }
    }
    
    return {
      success: true,
      student: student,
      fees: fees
    };
    
  } catch(error) {
    return { success: false, message: "Error: " + error.toString() };
  }
}

// Search for students
function searchStudent(query) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STUDENTS_SHEET);
  
  if (!sheet) {
    return { success: false, message: "Students sheet not found." };
  }
  
  if (!query || query.trim() === "") {
    return { success: false, message: "Please enter a search query" };
  }
  
  const data = sheet.getDataRange().getValues();
  const students = [];
  
  for (let i = 1; i < data.length; i++) {
    const name = data[i][1].toString().toLowerCase();
    const id = data[i][0].toString().toLowerCase();
    const queryLower = query.toLowerCase();
    
    if (name.includes(queryLower) || id.includes(queryLower)) {
      students.push({
        id: data[i][0],
        name: data[i][1],
        class: data[i][2],
        email: data[i][3]
      });
    }
  }
  
  if (students.length === 0) {
    return { success: false, message: "No students found matching your search" };
  }
  
  return { 
    success: true, 
    students: students 
  };
}

// Generate payment report for a term
function getReport(term) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentsSheet = ss.getSheetByName(PAYMENTS_SHEET);
  const feesSheet = ss.getSheetByName(FEES_SHEET);
  const studentsSheet = ss.getSheetByName(STUDENTS_SHEET);
  
  if (!paymentsSheet || !feesSheet || !studentsSheet) {
    return { success: false, message: "Required sheets not found." };
  }
  
  if (!term || term.trim() === "") {
    return { success: false, message: "Please enter a term" };
  }
  
  // Get all data
  const paymentData = paymentsSheet.getDataRange().getValues();
  const feeData = feesSheet.getDataRange().getValues();
  const studentData = studentsSheet.getDataRange().getValues();
  
  // Create lookup objects for fees and students
  const feeMap = {};
  for (let i = 1; i < feeData.length; i++) {
    feeMap[feeData[i][0]] = feeData[i][1]; // feeID -> term
  }
  
  const studentMap = {};
  for (let i = 1; i < studentData.length; i++) {
    studentMap[studentData[i][0]] = {
      name: studentData[i][1],
      class: studentData[i][2]
    };
  }
  
  // Filter payments for the specified term
  const payments = [];
  for (let i = 1; i < paymentData.length; i++) {
    const feeID = paymentData[i][2];
    const paymentTerm = feeMap[feeID];
    
    if (paymentTerm === term) {
      const studentID = paymentData[i][1];
      const studentInfo = studentMap[studentID];
      
      payments.push({
        paymentID: paymentData[i][0],
        studentID: studentID,
        name: studentInfo ? studentInfo.name : "Unknown",
        class: studentInfo ? studentInfo.class : "Unknown",
        amountPaid: paymentData[i][3],
        paymentDate: paymentData[i][4]
      });
    }
  }
  
  if (payments.length === 0) {
    return { success: false, message: "No payment records found for this term" };
  }
  
  return { 
    success: true, 
    term: term,
    payments: payments 
  };
}
