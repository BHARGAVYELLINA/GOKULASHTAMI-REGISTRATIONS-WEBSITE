function doPost(e) {
  var formData = e.postData.contents;
  var returnresult=parseFormData(formData,e.parameter.referringForm);
  
  return ContentService.createTextOutput(returnresult);
}

function parseFormData(dataString,eventtype) {
  var formDataArray = [];
  var formDataPairs = dataString.split('&');
  var currentData = {};
  var ss = SpreadsheetApp.openById('1PX85JfbKAsnLeNZxox6iolH5n7HZFgIHAQYxylAmn44');
  var debugsheet=ss.getSheetByName("check");
  var isIndividualEvent=["regarts.html","regmusic.html","regliter.html","regmedia.html","regdance.html","regcontribute.html"].includes(eventtype);

  for (var i = 0; i < formDataPairs.length; i++) {
    var keyValue = formDataPairs[i].split('=');
    var key = decodeURIComponent(keyValue[0]);
    var value = decodeURIComponent(keyValue[1].replace(/\+/g, ' '));

    if (!isIndividualEvent){
      if (key === 'teamName') {
        if (Object.keys(currentData).length > 0) {
          formDataArray.push(currentData);
          currentData = {};
        }
      }
    }
    else{
      if (key === 'fullName') {
        if (Object.keys(currentData).length > 0) {
          formDataArray.push(currentData);
          currentData = {};
        }
      }
    }
    // For properties with dots, use bracket notation
    if (key.includes('.')) {
      var parts = key.split('.');
      currentData[parts[0]] = currentData[parts[0]] || {};
      currentData[parts[0]][parts[1]] = value;
    } else {
      currentData[key] = value;
    }
  }

  if (Object.keys(currentData).length > 0) {
    formDataArray.push(currentData);
  }

  if(!isIndividualEvent){
    var referenceDepartment = formDataArray[1].department;
    var referenceCategory = formDataArray[1].category;
    var referenceTeamName = formDataArray[1].teamName;
    
    // // var referenceDepartment = currentData.department;
    // // var referenceCategory = currentData.category;
    // // var referenceTeamName = currentData.teamName;
    var uniqueFullNames = new Set();
    var uniqueRollNumbers = new Set();
    var uniquePhoneNumbers = new Set();

    // var referenceDepartment = currentData.department;
    // var referenceCategory = currentData.category;
    // var referenceTeamName = currentData.teamName;

    // Initialize reference values with the first member's data
    // var referenceDepartment = formDataArray[0].department;
    // var referenceCategory = formDataArray[0].category;
    // var referenceTeamName = formDataArray[0].teamName;

    for (var i = 1; i < formDataArray.length; i++) {
       var member = formDataArray[i];

    //   Debugging: Append reference and member values to debugsheet
    //   debugsheet.appendRow(["Reference Department:", referenceDepartment, "Member Department:", member.department]);
    //   debugsheet.appendRow(["Reference Category:", referenceCategory, "Member Category:", member.category]);
    //   debugsheet.appendRow(["Reference Team Name:", referenceTeamName, "Member Team Name:", member.teamName]);

    // Check if the department, category, and team name are NOT the same for all members
       if (member.department !== referenceDepartment) {
         return "Department is not the same for all team members.";
       }
       if (member.category !== referenceCategory) {
         return "Category is not the same for all members.";
       }
       if (member.teamName !== referenceTeamName) {
         return "Team name is not the same for all members.";
       }
        // Check if the full name is unique within the team
        if (uniqueFullNames.has(member.fullName)) {
          return "Full name of each team member must be unique.";
        }
        uniqueFullNames.add(member.fullName);

      // Check if the roll number matches the specified format
      // if (!rollNumberRegex.test(member.roll_no)) {
      //   return "Roll number format is not valid.";
      // }

        // Check if the roll number is unique within the team
        if (uniqueRollNumbers.has(member.roll_no)) {
          return "Roll number of each team member must be unique.";
        }
        uniqueRollNumbers.add(member.roll_no);

        // Check if the phone number is unique within the team
        if (uniquePhoneNumbers.has(member.phonenumber)) {
          return "Phone number of each team member must be unique.";
        }
        uniquePhoneNumbers.add(member.phonenumber);

      // for (var i = 0; i < formDataArray.length; i++) {
      //   debugsheet.appendRow(["Loop iteration",i]);
      //   var member = formDataArray[i]; // Get the current member object
      //   debugsheet.appendRow(["insideanother"],member.teamName,member.department,member.category);

      var memberIsRegistered = isMemberRegistered(ss, member.category, member.roll_no);
      if (memberIsRegistered) {
        return "One of the team members has already registered for this event in another team.";
      }
    }
    var teamNameIsUnique = isTeamNameUnique(ss, referenceCategory, referenceTeamName);
    if (!teamNameIsUnique) {
      return "Team name is already taken for this event.";
    }
    var limitedgroupevents=["GROUP DANCE"].includes(currentData.category);
    var sheetname="group events";
    var updatesheet="limited group events";
    var mainsheet=ss.getSheetByName(currentData.category);
    var sheet = ss.getSheetByName(sheetname);
    if(limitedgroupevents){
      var integratedvar=["INTEGRATED MSC-MATHS","INTEGRATED MSC-PHY","INTEGRATED MSC-CHEM","INTEGRATED MSC-DS","BSC-FSN","BA-ENG"].includes(currentData.department);
      if(integratedvar){
        var availableSlots=getAvailableSlotsForEvent(updatesheet,"INTEGRATED", currentData.category);
      }
      else{
        var availableSlots=getAvailableSlotsForEvent(updatesheet,currentData.department, currentData.category);
      }
      if(availableSlots>0){
        for (var i = 0; i < formDataArray.length; i++) {
          var member = formDataArray[i]; // Get the current member object
          // Access properties using dot notation
          var teamname=member.teamName;
          var fullName = member.fullName;
          var rollNumber = member.roll_no;
          var department = member.department;
          var phoneNumber = member.phonenumber;
          var gender = member.gender;
          var category = member.category;
          registrationvalues=[teamname,fullName,rollNumber,department,phoneNumber,gender,category];
          sheet.appendRow(registrationvalues);
          mainsheet.appendRow(registrationvalues);
        }
        if(integratedvar){
          updateAvailableSlotsInLimitedEventsSheet(updatesheet,"INTEGRATED", currentData.category, availableSlots - 1);
        }
        else{
          updateAvailableSlotsInLimitedEventsSheet(updatesheet,currentData.department, currentData.category, availableSlots - 1);
        }
        return "REGISTRATION SUCCESSFUL";
      }
      else{
        return "REGISTRATIONS CLOSED FOR THIS EVENT FOR YOUR DEPARTMENT";
      }
    }
    else{
      for (var i = 0; i < formDataArray.length; i++) {
        var member = formDataArray[i]; // Get the current member object

        // Access properties using dot notation
        var teamname=member.teamName;
        var fullName = member.fullName;
        var rollNumber = member.roll_no; // Use bracket notation for properties with dots
        var department = member.department;
        var phoneNumber = member.phonenumber;
        var gender = member.gender;
        var category = member.category;
        registrationvalues=[teamname,fullName,rollNumber,department,phoneNumber,gender,category];
        sheet.appendRow(registrationvalues);
        mainsheet.appendRow(registrationvalues);
      }
      return "REGISTRATION SUCCESSFUL";
    }
  }

  else{
    var studentIsRegistered = isStudentRegistered(ss, currentData.roll_no, currentData.category);

    if (studentIsRegistered) {
      return "You have already registered for this event.";
    }
    for (var i = 0; i < formDataArray.length; i++) {
      var member = formDataArray[i]; // Get the current member object
    }
    var mainsheet=ss.getSheetByName(currentData.category);
    var limitedindividualevents=["DANCE - SOLO",	"CREATIVE WRITING - ENGLISH",	"CREATIVE WRITING - TAMIL",	"CREATIVE WRITING - MALAYALAM",	"CREATIVE WRITING - HINDI", "CREATIVE WRITING - TELUGU"].includes(currentData.category);
    var sheetname='individual events';
    var sheet = ss.getSheetByName(sheetname);
    var updatesheet="limited individual events";
    if(limitedindividualevents){
      var integratedvar=["INTEGRATED MSC-MATHS","INTEGRATED MSC-PHY","INTEGRATED MSC-CHEM","INTEGRATED MSC-DS","BSC-FSN","BA-ENG"].includes(currentData.department);
      if(integratedvar){
        var availableSlots=getAvailableSlotsForEvent(updatesheet,"INTEGRATED", currentData.category);
      }
      else{
        var availableSlots=getAvailableSlotsForEvent(updatesheet,currentData.department, currentData.category);
      }
      if(availableSlots>0){
        for (var i = 0; i < formDataArray.length; i++) {
          var member = formDataArray[i]; // Get the current member object

          // Access properties using dot notation
          var fullName = member.fullName;
          var rollNumber = member.roll_no; // Use bracket notation for properties with dots
          var department = member.department;
          var phoneNumber = member.phonenumber;
          var gender = member.gender;
          var category = member.category;
          registrationvalues=[fullName,rollNumber,department,phoneNumber,gender,category];
          sheet.appendRow(registrationvalues);
          mainsheet.appendRow(registrationvalues);
        }
        if(integratedvar){
          updateAvailableSlotsInLimitedEventsSheet(updatesheet,"INTEGRATED", currentData.category, availableSlots - 1);
        }
        else{
          updateAvailableSlotsInLimitedEventsSheet(updatesheet,currentData.department, currentData.category, availableSlots - 1);
        }
        return "REGISTRATION SUCCESSFUL";
      }
      else{
        return "REGISTRATIONS CLOSED FOR THIS EVENT FOR YOUR DEPARTMENT";
      }
    }
    else{
      for (var i = 0; i < formDataArray.length; i++) {
        var member = formDataArray[i]; // Get the current member object

        // Access properties using dot notation
        var fullName = member.fullName;
        var rollNumber = member.roll_no; // Use bracket notation for properties with dots
        var department = member.department;
        var phoneNumber = member.phonenumber;
        var gender = member.gender;
        var category = member.category;
        registrationvalues=[fullName,rollNumber,department,phoneNumber,gender,category];
        sheet.appendRow(registrationvalues);
        mainsheet.appendRow(registrationvalues);
      }
      return "REGISTRATION SUCCESSFUL";
    }
  }
}

function isTeamNameUnique(ss, eventtype, teamName) {
  // Assuming you have separate sheets for each event category
  var eventSheet = ss.getSheetByName(eventtype);

  if (eventSheet) {
    var data = eventSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === teamName) { // Assuming team name is in the first column
        return false; // Team name is already taken for this event category
      }
    }
  }

  return true; // Team name is unique for this event category
}

function isMemberRegistered(ss, eventtype, roll_no) {
  // Assuming you have separate sheets for each event category
  var eventSheet = ss.getSheetByName(eventtype);

  if (eventSheet) {
    var data = eventSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === roll_no) { // Assuming roll number is in the third column
        return true; // Member is already registered for this event category in another team
      }
    }
  }

  return false; // Member is not registered for this event category in another team
}

function isStudentRegistered(ss, rollNumber, eventtype) {
  // Assuming you have separate sheets for each event category
  var eventSheet = ss.getSheetByName(eventtype);

  if (eventSheet) {
    var data = eventSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === rollNumber) { // Assuming roll number is in the second column
        return true; // Student with the same roll number is already registered for this event category
      }
    }
  }

  return false; // Student with the same roll number is not registered for this event category
}

function getAvailableSlotsForEvent(sheetname,department, event) {
  var ss = SpreadsheetApp.openById('1PX85JfbKAsnLeNZxox6iolH5n7HZFgIHAQYxylAmn44');
  var sheet = ss.getSheetByName(sheetname);
  var data = sheet.getDataRange().getValues();

  // Find the row index of the specified department in the first column
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === department) {
      // Find the column index of the specified event in the header row
      for (var j = 1; j < data[0].length; j++) {
        if (data[0][j] === event) {
          return data[i][j]; // Return available slots from the intersection
        }
      }
      break; // Department found, no need to search further
    }
  }

  return ContentService.createTextOutput('Department or Event not found ERROR.');
}

function updateAvailableSlotsInLimitedEventsSheet(sheetname, department, event, newAvailableSlots) {
  var ss = SpreadsheetApp.openById('1PX85JfbKAsnLeNZxox6iolH5n7HZFgIHAQYxylAmn44');
  var sheet = ss.getSheetByName(sheetname);
  var data = sheet.getDataRange().getValues();

  // Find the row index of the specified department in the first column
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === department) {
      // Find the column index of the specified event in the header row
      for (var j = 1; j < data[0].length; j++) {
        if (data[0][j] === event) {
          // Update the available slots for the specified event
          sheet.getRange(i + 1, j + 1).setValue(newAvailableSlots);
          
          // Check if the event is in the special list
          var specialEvents = ["CREATIVE WRITING - TAMIL", "CREATIVE WRITING - MALAYALAM", "CREATIVE WRITING - HINDI", "CREATIVE WRITING - TELUGU"];
          if (specialEvents.includes(event)) {
            // Reduce available slots for all special events in the list
            for (var k = 1; k < data[0].length; k++) {
              if (specialEvents.includes(data[0][k])) {
                sheet.getRange(i + 1, k + 1).setValue(newAvailableSlots);
              }
            }
          }
          break;
        }
      }
      break; // Department found, no need to search further
    }
  }
}

function sendMailToIndividualEventsManagersWithPDF() {
  var ss = SpreadsheetApp.openById('1PX85JfbKAsnLeNZxox6iolH5n7HZFgIHAQYxylAmn44');
  var outputSheet = ss.getSheetByName("individual events");
  var managerSheet = ss.getSheetByName("managers");
  
    /// Get data from the output and manager sheets
  var outputData = outputSheet.getDataRange().getValues();
  var managerData = managerSheet.getDataRange().getValues();
  
  // Create a mapping of categories to events and emails
  var categoryToEvent = {};
  
  // Create a mapping of events to manager emails
  var eventToManagers = {};
  
  // Loop through the manager data and create the mapping
  for (var i = 1; i < managerData.length; i++) {
    var event = managerData[i][0]; // Get the event name from the "Event" column (index 0)
    var email = managerData[i][2]; // Get the email address from the "Email" column (index 2)
    categoryToEvent[event] = email;
    
    if (!eventToManagers[event]) {
      eventToManagers[event] = [];
    }
    
    eventToManagers[event].push(email);
  }
  
  // Create an object to group participants by event
  var participantsByEvent = {};
  
  // Loop through the output data and group participants by event
  for (var j = 1; j < outputData.length; j++) {
    var fullName = outputData[j][0];
    var rollNumber = outputData[j][1];
    var department=outputData[j][2];
    var phoneNumber = outputData[j][3];
    var gender = outputData[j][4];
    var category = outputData[j][5];
    
    if (categoryToEvent.hasOwnProperty(category)) {
      var eventEmails = eventToManagers[category];
      
      if (!participantsByEvent[category]) {
        participantsByEvent[category] = [];
      }
      
      participantsByEvent[category].push([
        fullName,
        rollNumber,
        department,
        phoneNumber,
        gender,
        category
      ]);
    }
  }
  
  for (var eventCategory in participantsByEvent) {
    if (participantsByEvent.hasOwnProperty(eventCategory)) {
      var participants = participantsByEvent[eventCategory];
      
      // Retrieve event manager's email from the "managers" sheet
      var eventManagerEmail = getEventManagerEmail(managerSheet, eventCategory);
      
      if (eventManagerEmail) {
        // Create PDF content using the built-in PDF service
        var pdfBlob = createPDFContentIndividual(eventCategory, participants);
        
        // Include the checksum in the email subject
        var emailSubject = "Registration Details for Event: " + eventCategory;
        
        // Send email with PDF attachment to the event manager's email
        MailApp.sendEmail({
          to: eventManagerEmail,
          subject: emailSubject,
          body: 'Please find the attached PDF with registration details.',
          attachments: [pdfBlob]
        });
      }
    }
  }
}

function getEventManagerEmail(managerSheet, eventCategory) {
  var managerData = managerSheet.getDataRange().getValues();
  
  for (var i = 1; i < managerData.length; i++) {
    if (managerData[i][0] === eventCategory) {
      return managerData[i][2]; // Return the email from the "Email" column
    }
  }
  
  return null; // Return null if event category not found
}

function createPDFContentIndividual(eventCategory, participants) {
  // Create a new PDF document
  var pdf = DocumentApp.create(eventCategory);
  var body = pdf.getBody();
  
  // Add content to the PDF document
  body.appendParagraph("Event: " + eventCategory);
  body.appendParagraph(""); // Empty line
  
  // Create a table for participants
  var table = body.appendTable();  // Create an empty table
  
  // Add header row to the table
  var headerRow = table.appendTableRow();
  headerRow.appendTableCell("FULL NAME");
  headerRow.appendTableCell("ROLL NUMBER");
  headerRow.appendTableCell("DEPARTMENT");
  headerRow.appendTableCell("PHONE NUMBER");
  headerRow.appendTableCell("GENDER");
  headerRow.appendTableCell("CATEGORY");
  
  // Add participant data to the table
  for (var i = 0; i < participants.length; i++) {
    var participant = participants[i];
    var row = table.appendTableRow();
    
    // Append cells with participant data
    for (var j = 0; j < participant.length; j++) {
      row.appendTableCell(participant[j]);
    }
  }
  
  // Save and close the PDF document
  pdf.saveAndClose();
  
  // Get the PDF blob
  var pdfBlob = pdf.getAs("application/pdf");
  
  return pdfBlob;
}

function sendMailToGroupEventsManagersWithPDF() {
  var ss = SpreadsheetApp.openById('1PX85JfbKAsnLeNZxox6iolH5n7HZFgIHAQYxylAmn44');
  var outputSheet = ss.getSheetByName("group events");
  var managerSheet = ss.getSheetByName("managers");
  
    /// Get data from the output and manager sheets
  var outputData = outputSheet.getDataRange().getValues();
  var managerData = managerSheet.getDataRange().getValues();
  
  // Create a mapping of categories to events and emails
  var categoryToEvent = {};
  
  // Create a mapping of events to manager emails
  var eventToManagers = {};
  
  // Loop through the manager data and create the mapping
  for (var i = 1; i < managerData.length; i++) {
    var event = managerData[i][0]; // Get the event name from the "Event" column (index 0)
    var email = managerData[i][2]; // Get the email address from the "Email" column (index 2)
    categoryToEvent[event] = email;
    
    if (!eventToManagers[event]) {
      eventToManagers[event] = [];
    }
    
    eventToManagers[event].push(email);
  }
  
  // Create an object to group participants by event
  var participantsByEvent = {};
  
  // Loop through the output data and group participants by event
  for (var j = 1; j < outputData.length; j++) {
    var teamName = outputData[j][0];
    var fullName = outputData[j][1];
    var rollNumber = outputData[j][2];
    var department=outputData[j][3];
    var phoneNumber = outputData[j][4];
    var gender = outputData[j][5];
    var category = outputData[j][6];
    
    if (categoryToEvent.hasOwnProperty(category)) {
      var eventEmails = eventToManagers[category];
      
      if (!participantsByEvent[category]) {
        participantsByEvent[category] = [];
      }
      
      participantsByEvent[category].push([
        teamName,
        fullName,
        rollNumber,
        department,
        phoneNumber,
        gender,
        category
      ]);
    }
  }
  
  for (var eventCategory in participantsByEvent) {
    if (participantsByEvent.hasOwnProperty(eventCategory)) {
      var participants = participantsByEvent[eventCategory];
      
      // Retrieve event manager's email from the "managers" sheet
      var eventManagerEmail = getEventManagerEmail(managerSheet, eventCategory);
      
      if (eventManagerEmail) {
        // Create PDF content using the built-in PDF service
        var pdfBlob = createPDFContentGroup(eventCategory, participants);
        
        // Include the checksum in the email subject
        var emailSubject = "Registration Details for Event: " + eventCategory;
        
        // Send email with PDF attachment to the event manager's email
        MailApp.sendEmail({
          to: eventManagerEmail,
          subject: emailSubject,
          body: 'Please find the attached PDF with registration details.',
          attachments: [pdfBlob]
        });
      }
    }
  }
}

function createPDFContentGroup(eventCategory, participants) {
  // Create a new PDF document
  var pdf = DocumentApp.create(eventCategory);
  var body = pdf.getBody();
  
  // Add content to the PDF document
  body.appendParagraph("Event: " + eventCategory);
  body.appendParagraph(""); // Empty line
  
  // Create a table for participants
  var table = body.appendTable();  // Create an empty table
  
  // Add header row to the table
  var headerRow = table.appendTableRow();
  headerRow.appendTableCell("TEAM NAME");
  headerRow.appendTableCell("FULL NAME");
  headerRow.appendTableCell("ROLL NUMBER");
  headerRow.appendTableCell("DEPARTMENT");
  headerRow.appendTableCell("PHONE NUMBER");
  headerRow.appendTableCell("GENDER");
  headerRow.appendTableCell("CATEGORY");
  
  // Add participant data to the table
  for (var i = 0; i < participants.length; i++) {
    var participant = participants[i];
    var row = table.appendTableRow();
    
    // Append cells with participant data
    for (var j = 0; j < participant.length; j++) {
      row.appendTableCell(participant[j]);
    }
  }
  
  // Save and close the PDF document
  pdf.saveAndClose();
  
  // Get the PDF blob
  var pdfBlob = pdf.getAs("application/pdf");
  
  return pdfBlob;
}