(function () {
  'use strict';

  // DEMO ONLY: credentials are checked in the browser, so anyone can read them.
  // Replace checkCredentials() with a call to your server for real use.
  var STAFF_ID_PATTERN = /^[A-Z]{3}\d{4}$/;   // e.g. NRB1234
  var DEMO_PASSWORD = '1234';
  var HOME_PAGE = 'hackathon_v4-main/index.html';         // page to open after login

  var form = document.getElementById('login-form');
  var idInput = document.getElementById('employeeId');
  var pwInput = document.getElementById('password');
  var errorBox = document.getElementById('error');
  



  function showError(message, field) {
    errorBox.textContent = message;
    errorBox.hidden = false;
    idInput.classList.remove('invalid');
    pwInput.classList.remove('invalid');
    if (field) { field.classList.add('invalid'); field.focus(); }
  }

  function clearError() {
    errorBox.hidden = true;
    errorBox.textContent = '';
    idInput.classList.remove('invalid');
    pwInput.classList.remove('invalid');
  }

  function checkCredentials(id, password) {
    return STAFF_ID_PATTERN.test(id) && password === DEMO_PASSWORD;
  }

  form.addEventListener('submit', function (e) {
    e.preventDefault();
    clearError();

    var id = idInput.value.trim().toUpperCase();
    var password = pwInput.value;

    if (!id) { return showError('Employee ID cannot be blank.', idInput); }
    if (!password) { return showError('Password cannot be blank.', pwInput); }
    if (!STAFF_ID_PATTERN.test(id)) {
      return showError('Employee ID must be 3 letters followed by 4 digits, e.g. NRB1234.', idInput);
    }
    if (!checkCredentials(id, password)) {
      return showError('Incorrect Employee ID or password.', pwInput);
    }

    try { sessionStorage.setItem('staffId', id); } catch (err) {}
    window.location.href = HOME_PAGE;
  });

  [idInput, pwInput].forEach(function (el) {
    el.addEventListener('input', clearError);
  });
})();
