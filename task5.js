"use strict";
const organizationURI = "https://org94d5d311.crm11.dynamics.com"; //The URL to connect to CRM Online
const tenant = "qSolutions184.onmicrosoft.com"; //The name of the Azure AD organization you use
const clientId = "da523663-87f5-4e35-9fa7-40053d10c8e9"; //The ClientId you got when you registered the application
const pageUrl = "http://localhost:5501/task5/task5.html"; //The URL of this page in your development environment when debugging.
let user,
  authContext,
  loginButton,
  logoutButton,
  startBtn,
  cancelBtn,
  startDateInput,
  endDateInput,
  startDate,
  endDate,
  progressBar,
  progressS,
  progressAll,
  statusMessage,
  birthdayArticle;
const contactsQuery =
  "/api/data/v8.0/contacts?$select=fullname,birthdate,emailaddress1,cr7c6_dateofthelastbirthdaycongratulations";
const apiVersion = "/api/data/v8.0/";
//Configuration data for AuthenticationContext
const endpoints = {
  orgUri: organizationURI,
};

let isCanceled = false;

window.config = {
  tenant: tenant,
  clientId: clientId,
  postLogoutRedirectUri: pageUrl,
  endpoints: endpoints,
  cacheLocation: "localStorage", // enable this for IE, as sessionStorage does not work for localhost.
};

document.onreadystatechange = function () {
  if (document.readyState == "complete") {
    startBtn = document.querySelector(".btn-start");
    cancelBtn = document.querySelector(".btn-cancel");
    startDateInput = document.querySelector(".date-start");
    endDateInput = document.querySelector(".date-end");
    progressBar = document.querySelector(".progress-bar");
    progressS = document.querySelector(".progress-bar__text span:first-child");
    progressAll = document.querySelector(".progress-bar__text span:last-child");
    statusMessage = document.querySelector(".status__message");
    birthdayArticle = document.querySelector(".birthday-article");

    //Set DOM elements referenced by scripts
    loginButton = document.getElementById("login");
    logoutButton = document.getElementById("logout");
    //Event handlers on DOM elements

    startBtn.addEventListener("click", StartHandler);
    cancelBtn.addEventListener("click", CancelHandler);

    loginButton.addEventListener("click", login);
    logoutButton.addEventListener("click", logout);

    authenticate();

    if (user) {
      loginButton.style.display = "none";
      logoutButton.style.display = "block";
      birthdayArticle.style.display = "block";
    } else {
      loginButton.style.display = "block";
      logoutButton.style.display = "none";
      birthdayArticle.style.display = "none";
    }
  }
};

// Function that manages authentication
function authenticate() {
  //OAuth context
  authContext = new AuthenticationContext(config);
  // Check For & Handle Redirect From AAD After Login
  const isCallback = authContext.isCallback(window.location.hash);
  if (isCallback) {
    authContext.handleWindowCallback();
  }
  const loginError = authContext.getLoginError();
  if (isCallback && !loginError) {
    window.location = authContext._getItem(
      authContext.CONSTANTS.STORAGE.LOGIN_REQUEST
    );
  } else {
    console.error(loginError);
  }
  user = authContext.getCachedUser();
}

//function that logs in the user
function login() {
  authContext.login();
}
//function that logs out the user
function logout() {
  authContext.logOut();
  birthdayArticle.style.display = "none";
}

function StartHandler() {
  isCanceled = false;
  ClearData();

  startDate = new Date(startDateInput.value);
  endDate = new Date(endDateInput.value);
  const nowDate = new Date();

  if (
    DateToStr(startDate) !== DateToStr(nowDate) ||
    isNaN(endDate) ||
    endDate - startDate >= 604800000 ||
    endDate - startDate < 0
  ) {
    alert(
      "The start date is not earlier than today’s date and not later than today’s date, and the end date is no more than 7 days from today’s date"
    );
    return;
  }

  getContacts();
}

async function SentBirthdayMessage(contacts) {
  progressAll.innerText = contacts.length;
  let progressBarUpdeter = ProgressUpdate();

  for (let i = 0; i < contacts.length; i++) {
    if (contacts[i].emailaddress1 === null) {
      progressBarUpdeter(false);
      continue;
    }
    if (isCanceled) {
      break;
    } else {
      await Email.send({
        Host: "sandbox.smtp.mailtrap.io",
        Username: "8b0142f1413bb0",
        Password: "bb2206b198685a",
        To: contacts[i].emailaddress1,
        From: "qTU@qSolutions184.onmicrosoft.com",
        Subject: "Birthday Email",
        Body: "Happy Birthday dear " + contacts[i].fullname + "!",
      }).then(
        function (response) {
          authContext.acquireToken(organizationURI, function (error, token) {
            if (error || !token || isCanceled) {
              progressBarUpdeter(false);
              console.error("Error: ", error);
              return;
            }
            updateContact(contacts[i].contactid, token).then((response) => {
              if (!response.ok) {
                progressBarUpdeter(false);
                console.error("Error", response);
                return response;
              }
              progressBarUpdeter(true);
            });
          });
        },
        function (error) {
          console.error(error);
        }
      );
    }
  }
}

// GET CONTACTS
async function getContacts() {
  await authContext.acquireToken(organizationURI, retrieveContacts);
}

// RETRIEVE CONTACTS
function retrieveContacts(error, token) {
  // Handle ADAL Errors.
  if (error || !token) {
    console.error("ADAL error occurred: " + error);
    return;
  }

  const url = encodeURI(organizationURI + contactsQuery);

  const contactsData = fetch(url, {
    method: "GET",
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/json",
      "Content-Type": "application/json; charset=utf-8",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
    },
  })
    .then((response) => {
      return response.json();
    })
    .then((data) => {
      const contacts = data.value;
      const contactsFiltered = FilterEntities(contacts);

      SentBirthdayMessage(contactsFiltered);
    });

  return contactsData;
}

function CancelHandler() {
  isCanceled = confirm("Are you sure that you want to cancel the processing?");
}

function ClearData() {
  progressAll.innerText = "0";
  progressS.innerText = "0";
  progressBar.value = 0;
  statusMessage.innerText = "";
}

async function updateContact(contactId, token) {
  const data = JSON.stringify({
    cr7c6_dateofthelastbirthdaycongratulations: DateToStr(new Date()),
  });

  const url = encodeURI(
    organizationURI + `/api/data/v8.0/contacts(${contactId})`
  );

  const patchContactData = await fetch(url, {
    method: "PATCH",
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/json",
      "Content-Type": "application/json; charset=utf-8",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
    },
    body: data,
  });

  return patchContactData;
}

// Progress updeter
function ProgressUpdate() {
  let progressValue = 0;

  return function (bool) {
    progressS.innerText = +(+progressS.innerText || 0) + +bool;
    progressValue++;
    if (+progressS.innerText > 0 && +progressAll.innerText > 0) {
      progressBar.value = (progressValue / progressAll.innerText) * 100;
    }

    if (progressValue === +progressAll.innerText) {
      statusMessage.innerText =
        "Processing is cancelled successfully. " +
        (+progressAll.innerText || 0) +
        " contacts are processed, " +
        (+progressS.innerText || 0) +
        " congratulation e-mails were suceesfully sent (sending of " +
        ((+progressAll.innerText || 0) - (+progressS.innerText || 0)) +
        " e-mails failed)";
    }
  };
}

function FilterEntities(entities) {
  const NewArray = [];

  for (let i = 0; i < entities.length; i++) {
    if (
      CheckBirthday(startDate, endDate, entities[i].birthdate) &&
      entities[i].cr7c6_dateofthelastbirthdaycongratulations === null
    ) {
      NewArray.push(entities[i]);
    }
  }

  return NewArray;
}

function CheckBirthday(start, end, date) {
  const birthday = new Date(date);
  birthday.setFullYear(start.getFullYear());
  if (birthday >= start && birthday <= end) {
    return true;
  }
  return false;
}

function DateToStr(date) {
  return `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`;
}
