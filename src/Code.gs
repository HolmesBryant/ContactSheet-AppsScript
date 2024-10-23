/**
 * @OnlyCurrentDoc
 */

/**
 * @author Holmes Bryant <https://github.com/HolmesBryant>
 * @license MIT
 * @see {@link https://holmesbryant.github.io/ContactSheet-AppsScript/}
 */

var globals = PropertiesService.getScriptProperties();

/**
 * Your apps script url. Paste it here after you deploy the script, then save. This a convenient way to remember the url
 * It is displayed in the UI sidebar when you choose "ContactSheet > Options".
 *
 * @type {string}
 */
const appUrl = "";

/**
 * The name of the sheet in which you want to store the results
 *
 * @type {string}
 */
var sheetName = globals.getProperty('sheetName') || "Sheet1";

/**
 * The email address to which you want to send notifications of form submissions. If you do not want to send notifications, set this to an empty string.
 * It is highly recommended to leave this as the default setting. See the README.
 * @type {string}
 */
var sendTo = Session.getEffectiveUser().getEmail();

/**
 * The message to send to the client upon a successful form submission.
 *
 * @type {string}
 */
var successMessage = globals.getProperty('successMessage') || "Form submitted successfully!";

/**
 * The message to send to the client when an error occurs and "debug" is falsy
 *
 * @type {string}
 */
var errorMessage = globals.getProperty('errorMessage') || "There was a problem processing the form. The body was consumed but it could not be fully digested";

/**
 * The email address of the Google Group to which you want to add contacts who wish to be on your mailing list.
 * If you do not want to add the contact to a group, set this to an empty string.
 *
 * @type {string}
 */
var groupEmail = globals.getProperty('groupEmail') || "";

/**
 * Whether to add the person submitting the form to your Google Contacts.
 * If you want to add the person, set this to "yes".
 * If you do not want to add the person, set this to an empty string.
 *
 * @type {"yes"|""}
 */
var addContact = globals.getProperty('addContact') || "";

/**
 * Whether to send discrete error messages back to the client which submitted the form when errors occur.
 * Set to "yes" to send error messages.
 * @type {"yes"|""}
 */
var debug = globals.getProperty('debug') || "";

/**
 * Where the data [or contact] came from. If you use the Sheet to store contacts encountered from souces other than the contact form, this helps to identify how the contact entered your system.
 * @type {String}
 */
var dataSource = "Contact Form";

/**
 * The name of the app. Only used for display purposes.
 * You can change it if your UI requires specific branding.
 * @type {string}
 */
var appName = "ContactSheet";

/**
 * Test data to use when testing the script with [run > testPost]
 * This mimics the object created from data submitted by an HTML form.
 * Change email to a valid email address.
 *
 * @type {Object}
 */
var testData = {
  firstName: "Homie",
  lastName: "deClown",
  email: "valid@email.address",
  phone: "1234567890",
  subject: "Test Subject",
  message: "This is a test message.",
  mailingList: "on"
}

/* DO NOT EDIT BELOW THIS LINE */

var errors = [];

/**
 * Display the README file (in the sheet window)
 * Uses the html file "readme.html"
 */
  function showReadme() {
    var output = HtmlService.createHtmlOutputFromFile('readme');
    SpreadsheetApp.getUi().showModalDialog(output, appName);
  }

/**
 * Build and display the sidebar in the sheet
 * Uses the html template "sidebar.html"
 */
function showSidebar() {
  let html = HtmlService.createTemplateFromFile('sidebar')
    .evaluate()
    .setTitle(appName);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Save the updated values from the Options sidebar in the sheet UI
 *
 * @param {object} obj - A Google generated object which is passed to the function.
 */
function saveVariables(obj) {
  try {
    PropertiesService.getScriptProperties().setProperties(obj);
    SpreadsheetApp.getUi().alert("Options Saved!");
    return true;
  } catch (error) {
    console.log('Error saving: ' + error);
    return false;
  }
}

/**
 * Add a custom menu item
 */
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu(appName)
  .addItem('README', 'showReadme')
  .addItem('Options', 'showSidebar')
  .addToUi();
}

/**
 * Add data to google sheet
 *
 * @param {string} firstName - The user's first name
 * @param {string} lastName - The user's  last name
 * @paran {string} email - The user's email
 * @param {string} phone - The user's phone number
 * @param {"Yes"|"No"} mailingList - Whether the user wishes to be added to the google group (mailing list)
 * @param {string} source - Where the contact came from. The default is "Contact Form"
 * @param {string} subject - The subject of the message
 * @param {string} message - The message
 *
 * @throws {Error}
 */
function addToSheet(firstName, lastName, email, phone, mailingList, source, subject, message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(["Timestamp", "First Name", "Last Name", "Email", "Phone", "Mailing List", "Source", "Subject", "Message"]);
    }

    sheet.appendRow([new Date(), firstName, lastName, email, phone, mailingList, source, subject, message]);
    console.log("Sheet successfully set");
  } catch (error) {
    throw error;
  }
}

/**
 * Format a phone number
 * Replaces parenthese, dashes and dots with spaces
 *
 * @param {string} phoneNumber - The phone number to format
 *
 * @return {string} - The formatted phone number
 */
function formatPhoneNumber(phoneNumber) {
  phoneNumber = phoneNumber.replace(/[^\s\d+()\-]/g, " ");
  const re1 = /(\d{3})(\d{3})(\d{4})/; // US: no spaces, dots, dashes or parenthese. just one long number
  const re2 = /^(\+\d+)[ .-]?(\d+)$/; // Intl:  +country code followed by space|dot|dash followed by one long number
  const re3 = /^(\+\d{1,3})?[-. ]?\(?([\d]+)\)?[-. ]+(.+)/; // other formats
  if (phoneNumber.match(re1)) {
    phoneNumber = phoneNumber.replace(re1, "$1 $2 $3");
  } else if(phoneNumber.match(re2)) {
    phoneNumber = phoneNumber.replace(re2, "$1 $2");
  } else {
    phoneNumber = phoneNumber.replace(re3, "$1 $2 $3");
  }

  return phoneNumber;
}

/**
 * Check if Google Contact already exists.
 * Searches contacts by firstName and email
 *
 * @param {string} firstName - The given name of the contact to search
 * @param {string} email - The email of the contact to search
 *
 * @return {boolean}
 * @throws {Error}
 */
function doesContactExist(firstName, email) {
  let contacts = [];
  try {
    const results = People.People.Connections.list('people/me', { personFields: "emailAddresses,names" });
    if (results.connections) contacts = results.connections.map( item => ({name: (item.names)? item.names[0].givenName : "", emails: item.emailAddresses.map( email => email.value)}));
    const filtered = contacts.filter( item => item && item.name?.includes(firstName) && item.emails?.includes(email));

    if (filtered.length > 0) {
      console.log("Contact exists");
      return true;
    } else {
      console.log("Contact NOT found");
      return false;
    }
  } catch (error) {
    throw error;
  }
}

/**
 * Add contact to your Google Contacts
 *
 * @param firstName {string} - The first name
 * @param lastName {string} - The last name
 * @param email {string} - the email
 * @param phone {string|null} - The phone number
 *
 * @throws {Error}
 */
function addToContacts(firstName, lastName, email, phone) {
  const contact = {
    names: [{ givenName: firstName, familyName: lastName }],
    emailAddresses: [{ value: email }]
  };

  if (phone) {
    contact.phoneNumbers = [{ value: phone }];
  }

  try {
    People.People.createContact(contact);
    console.log("Contact created");
  } catch (error) {
    throw error;
  }
}

/**
 * Checks if a given email address corresponds to an existing Google Group.
 *
 * @param {string} groupEmail The email address to check (e.g., "mygroup@domain.com").
 *
 * @return {boolean} True if the group exists, false otherwise.
 */
function doesGroupExist(groupEmail) {
  try {
    GroupsApp.getGroupByEmail(groupEmail);
    console.log('Group exists: ' + groupEmail);
    return true;
  } catch (error) {
    console.log('Group not found:' + groupEmail);
    return false;
  }
}

/**
 * Send email inviting person to join a Google Group
 *
 * @param {string} userEmail - Email address which should receive the invitation to join.
 * @param {string} groupEmail - Email address associated with the the Google Group.
 *
 * @throws {Error}
 */
function inviteToGroup(userEmail, groupEmail) {
  const parts = groupEmail.split('@');
  const link = `https://groups.google.com/a/${parts[1]}/g/${parts[0]}`;
  const subject = `Invitation to join our mailing list: ${groupEmail}`;
  const groupLink = `<a href="${link}" style="padding:8px; background:#1e90ff; color:#ffffff; border-radius:5px; text-decoration:none; font-weight:bold; font-size:20px">Request to join our mailing list</a>`;
  const body = `
  <!DOCTYPE html>
  <html>
    <head>
      <meta charset="utf-8">
      <title>Invitation to join our mailing list</title>
    </head>
    <body style="padding:16px">
      <p>Thank you for your desire to join the mailing list at ${groupEmail}</p>
      <p>Please click the link below to go to our Google Groups mailing list.</p>
      <p>Then click on the button labeled "Ask to join group"</p>
      <p>If you did not request to join our mailing list, you can safely disregard this message.</p>
      <p>${groupLink}</p>
    </body>
  </html>`;

  try {
    MailApp.sendEmail({
      to: userEmail,
      subject: subject,
      htmlBody: body
    });
    console.log(`Group Invite Sent: To ${userEmail} Link: ${link}`);
  } catch(error) {
    throw error;
  }
}

/**
 * Sends an email notification when a new contact is added to the spreadsheet
 *
 * @param {string} sendTo - The email to which the notification should be sent
 * @param {string} firstName - The contact's first name
 * @param {string} lastName  - The contact's last name
 * @param {string} email - The contact's email
 * @param {string} phone - The contact's phone
 * @param {string} subject - The subject of the message
 * @param {string} message - The message the contact sent via the contact form
 * @param {Boolean} invitedToGroup - Whether the contact was sent an invitation to join the Google Group
 * @param {Boolean} contactExists - Whether the contact already exists in this account's Google Contacts
 *
 * @throws {Error}
 */
function sendNotification(sendTo, firstName, lastName, email, phone, subject, message, invitedToGroup, contactExists, addContact) {
  let contactsMsg = "";
  let contactsBtn = "";

  const phoneNumber = (phone)? `<p><b>Phone:</b> ${phone}</p>` : "";
  const invited = (invitedToGroup === true) ? `<p>This person was sent an invitation to join: ${groupEmail}</p>` : "";
  const replyBtn = `<a href="mailto:${email}?subject=Re: ${subject}" style="padding:8px; background:#1e90ff; color:#ffffff; border-radius:5px; text-decoration:none; font-weight:bold; font-size:20px">Reply</a>`;

  if (contactExists) {
    contactsMsg = "This person is already in your Contacts.";
    contactsBtn = `<a href="https://contacts.google.com/search/${encodeURIComponent(email)}" style="padding:8px; background:#1e90ff; color:#ffffff; border-radius:5px; text-decoration:none; font-weight:bold; font-size:20px">View Contact</a>`;
  } else if(addContact === "yes") {
    contactsMsg = "This person was added to your contacts";
    contactsBtn = `<a href="https://contacts.google.com/search/${encodeURIComponent(email)}" style="padding:8px; background:#1e90ff; color:#ffffff; border-radius:5px; text-decoration:none; font-weight:bold; font-size:20px">View Contact</a>`;
  }

  let body = `<!DOCTYPE html>
  <html>
    <head>
      <meta charset="utf-8">
      <title>New Message</title>
    </head>
    <body style="padding:16px">
      <p style="margin-bottom:20px">${replyBtn}</p>
      <p><b>From:</b> ${firstName} ${lastName} : ${email} </p>
      ${phoneNumber}
      <p><b>Subject:</b> ${subject}</p>
      <div>${message}</div>
      <hr>
      ${invited}
      <p>${contactsMsg}</p>
      <p>${contactsBtn}</p>
    </body>
  </html>`;

  try {
    MailApp.sendEmail({
      to: sendTo,
      subject: `Contact Form: ${subject}`,
      htmlBody: body
    });
    console.log("Notification sent to: " + sendTo);
  } catch (error) {
    throw error;
  }
}

/**
 * Send response back to the url which submitted the form
 *
 * @param {Boolean} logOutput - Whether to include the response in the Execution log
 * @param {Boolean} invitedToGroup - Whether the contact was sent an invitation to join the Google Group
 *
 * @return {ContentService} Outputs a response message
 */
function responseMessage(logOutput = false, invitedToGroup = false) {
  let response;

  if (errors.length > 0) {
    if (debug === "yes") {
      response = '<p><strong role="alert">Error!</strong></p>';
      response += errors.map( item => `<p>${item}</p>`).join("");
    } else {
      response = `<strong role="alert">${errorMessage}</strong>`;
    }
  } else {
    response = `<p><strong>${successMessage}</strong></p>`;
    if (invitedToGroup === true) {
      response += "<p>We have sent you a confirmation email inviting you to join our mailing list</p>";
    }
  }

  if (logOutput === true) {
    return response;
  } else {
    return ContentService.createTextOutput(response).setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * Used for testing if the script is live and accessible
 * Create a Test deployment and paste the resulting app url in the browser address bar
 *
 * @return {ContentService} Outputs a response message
 */
function doGet() {
  const response = "Contact Form Service Operational";
  return ContentService.createTextOutput(response).setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Test function for debugging before deployment
 * Creates a mock "e" object which is then passed to doPost.
 * To use, choose "testPost" from the dropdown at the top of the Apps Script editor window. Then click "Run".
 *
 * @return {ContentService} Outputs a response message
 */
function testPost() {
  const phone = testData.phone.split(" ").join("+");
  const subject = testData.subject.split(" ").join("+");
  const message = testData.message.split(" ").join("+");

  // Create a mock 'e' object
  let e = {
    // Add parameters as they would be submitted in a POST request
    parameter: {
      firstName: testData.firstName,
      lastName: testData.lastName,
      email: testData.email,
      phone: testData.phone,
      subject: testData.subject,
      message: testData.message,
      mailingList: testData.mailingList
    },
    postData: {
      // URL encoded string
      contents: `firstName=${testData.firstName}&lastName=${testData.lastName}&email=${testData.email}&phone=${phone}&subject=${subject}&message=${message}&mailingList=${testData.mailingList}`,
      type: "application/x-www-form-urlencoded"
    },

    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    }
  };

  return doPost(e);
}

/**
 * The entry point of the script. When you create the trigger, set it to execute this function.
 *
 * @param {Object} e - An event object created by Google and passed to the function when an http request is received
 *
 * @return {ContentService} Outputs a response message
 */
function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  var data = e.parameter;
  const timestamp = new Date();
  const firstName = data.firstName;
  const lastName = data.lastName;
  const email = data.email;
  let phone = data.phone;
  const subject = data.subject;
  const message = data.message;
  const mailingList = Boolean(data.mailingList) === true ? "Yes" : "No";
  const source = dataSource;
  // "mirage" is a honey pot input in the HTML form that gets removed when form is submitted on the client. See the README.
  const mirage = Boolean(data.mirage);

  // If a form input named "mirage" was sent with the form submission, consider it a bot attack
  if (mirage) {
    console.log("Honeypot detected. Suspected bot.");
    const honeyPotMsg = (debug === "yes")? "Suspected bot!" : "";
    return ContentService
    .createTextOutput(honeyPotMsg)
    .setMimeType(ContentService.MimeType.TEXT);
  }

  // Format phone number
  if (phone) {
    try {
      phone = formatPhoneNumber(phone);
    } catch (error) {
      errors.push("Phone format error: " + error);
      console.log("Phone format error: " + error);
    }
  }

  // Add data to spreadsheet
  try {
    sheet.appendRow([timestamp, firstName, lastName, email, phone, mailingList, source, subject, message]);
  } catch (error) {
    errors.push("Sheet Error: " + error);
    console.log("Sheet error: " + error);
  }

  // Check if person exists in Google Contacts
  try {
    var contactExists = doesContactExist(firstName, email);
  } catch (error) {
    errors.push("Error checking if contact exists: " + error);
    console.log("Error checking if contact exists: " + error);
  }

  // Add contact to Google Contacts
  if (!contactExists && addContact === "yes") {
    try {
      addToContacts(firstName, lastName, email, phone);
    } catch (error) {
      console.log("Error adding Contact: " + error);
      errors.push("Error adding Contact: " + error);
    }
  }

  let invited = false;
  let groupExists = false;

  // Send invite to join Google Group
  if (!!groupEmail && mailingList === "Yes")  {

    // Check if the google group exists
    try {
      groupExists = doesGroupExist(groupEmail);
    } catch(error) {
      errors.push('Error checking mailing list: ' + error);
      console.log('Error checking mailing list: ' + error);
    }

    if (groupExists) {
      try {
        inviteToGroup(email, groupEmail);
        invited = true;
      } catch (error) {
        console.log("Error sending group invite: " + error);
        errors.push("Error sending group invite: " + error);
      }
    }
  }

  // Send notification email
    try {
      console.log("Send Notification to: " + sendTo);
      sendNotification(sendTo, firstName, lastName, email, phone, subject, message, invited, contactExists, addContact);
    } catch (error) {
      console.log("Error sending notification: " + error);
      errors.push(`Error sending notification: ${error}`);
    }

  console.log(responseMessage(true, invited));
  return responseMessage(false, invited);
}
