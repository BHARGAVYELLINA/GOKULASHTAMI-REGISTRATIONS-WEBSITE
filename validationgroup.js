document.addEventListener('DOMContentLoaded', function () {
    const form = document.getElementById('registration-form');
  
    form.addEventListener('submit', function (event) {
      let isValid = true;
  
      // Clear previous error messages
      const errorDivs = document.querySelectorAll('.error');
      errorDivs.forEach(div => div.textContent = '');
  
      // Validate team name
      const teamNameInput = document.querySelector('input[name="teamName"]');
      if (!teamNameInput.value.trim()) {
        isValid = false;
        const nameErrorDiv = document.getElementById('teamerror');
        nameErrorDiv.textContent = 'Team name is required.';
        nameErrorDiv.classList.add('error-message'); // Add this line to apply CSS class
      }
      else{
        const nameErrorDiv = document.getElementById('teamerror');
        nameErrorDiv.textContent = ''; // Clear the error message
        nameErrorDiv.classList.remove('error-message'); // Remove the error class if present
      }

      // Validate full name
      const fullNameInput = document.querySelector('input[name="fullName"]');
      if (!fullNameInput.value.trim()) {
        isValid = false;
        const nameErrorDiv = document.getElementById('nameerror');
        nameErrorDiv.textContent = 'Full name is required.';
        nameErrorDiv.classList.add('error-message'); // Add this line to apply CSS class
      }
      else{
        const nameErrorDiv = document.getElementById('nameerror');
        nameErrorDiv.textContent = ''; // Clear the error message
        nameErrorDiv.classList.remove('error-message'); // Remove the error class if present
      }
  
      
      // Validate roll number
    const rollNumberInput = document.querySelector('input[name="roll_no"]');
    const rollNumberValue = rollNumberInput.value;
    const rollNumberRegex = /\bCB\.\w{2}\.U4(?:CSE|ECE|ELC|SC\.I5PHY|MEE|CCE)\d{5}\b/gi;
  
    if (!rollNumberRegex.test(rollNumberValue)) {
      isValid = false;
      const rollNumberErrorDiv = document.getElementById('rollnumbererror');
      rollNumberErrorDiv.textContent = 'Enter a valid roll number (e.g., CB.XX.U4XXXXXXXX).';
      rollNumberErrorDiv.classList.add('error-message');
    } else {
      const rollNumberErrorDiv = document.getElementById('rollnumbererror');
      rollNumberErrorDiv.textContent = ''; // Clear the error message
      rollNumberErrorDiv.classList.remove('error-message'); // Remove the error class if present
      
      const joiningYear = parseInt(rollNumberValue[11] + rollNumberValue[12]);
      
      if (joiningYear >= 23) {
        // Joining year is greater than or equal to 23
        alert('1st years are not allowed to participate');
        event.preventDefault(); // Prevent form submission
        return;
      }
      
      if (joiningYear <= 19) {
        // Joining year is less than or equal to 19
        alert('5th years are not allowed to participate');
        event.preventDefault(); // Prevent form submission
        return;
      }
    }
  
  
      // Validate phone number
      const phoneNumberInput = document.querySelector('input[name="phonenumber"]');
      const phoneNumberRegex = /^\d{10}$/;
      if (!phoneNumberRegex.test(phoneNumberInput.value)) {
        isValid = false;
        const phoneNumberErrorDiv = document.getElementById('numbererror');
        phoneNumberErrorDiv.textContent = 'Enter a valid 10-digit phone number.';
        phoneNumberErrorDiv.classList.add('error-message');
      }
      else{
        const nameErrorDiv = document.getElementById('numbererror');
        nameErrorDiv.textContent = ''; // Clear the error message
        nameErrorDiv.classList.remove('error-message'); // Remove the error class if present
      }
  
      // Validate department
      const departmentSelect = document.querySelector('select[name="department"]');
      if (departmentSelect.value === '') {
        isValid = false;
        const departmentErrorDiv = document.getElementById('departmenterror');
        departmentErrorDiv.textContent = 'Select a department.';
        departmentErrorDiv.classList.add('error-message');
      }
      else{
        const nameErrorDiv = document.getElementById('departmenterror');
        nameErrorDiv.textContent = ''; // Clear the error message
        nameErrorDiv.classList.remove('error-message'); // Remove the error class if present
      }
  
      // Validate gender
      const genderselect = document.querySelector('select[name="gender"]');
      if (genderselect.value === '') {
        isValid = false;
        const genderErrorDiv = document.getElementById('gendererror');
        genderErrorDiv.textContent = 'Select a category.';
        genderErrorDiv.classList.add('error-message');
      }
      else{
        const nameErrorDiv = document.getElementById('gendererror');
        nameErrorDiv.textContent = ''; // Clear the error message
        nameErrorDiv.classList.remove('error-message'); // Remove the error class if present
      }
  
      // Validate category
      const categorySelect = document.querySelector('select[name="category"]');
      if (categorySelect.value === '') {
        isValid = false;
        const categoryErrorDiv = document.getElementById('categoryerror');
        categoryErrorDiv.textContent = 'Select a category.';
        categoryErrorDiv.classList.add('error-message');
      }
      else{
        const nameErrorDiv = document.getElementById('categoryerror');
        nameErrorDiv.textContent = ''; // Clear the error message
        nameErrorDiv.classList.remove('error-message'); // Remove the error class if present
      }
  
      if (!isValid) {
        event.preventDefault(); // Prevent form submission if there are errors
      }
      
    });
  });