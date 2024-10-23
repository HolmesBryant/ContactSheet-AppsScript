# README #

## Version ##

* 1.0

## Overview ##

This Google Apps Script is meant to be bound to a Google Sheet. It should work in both Workspace accounts and free accounts. It receives contact data from an HTML form on your website and:

* Checks for the presence of an input named "mirage"
   - If it finds such an input, it assumes it is a bot trap and rejects the form submission.
* Checks for the presence of an input named "phone"
    - If it finds such an input, it tries to format the value to ` x(n) x(n) x(n) ...` so all phone numbers have consistent formatting. For example, "123 456 7890" or "+44 20 7946 0808".
* Uses the value of the email field and the firstName field to check if the person exists in your Google Contacts.
    - If the contact does not exist in Google Contacts AND the option (accessed from ContactSheet > Options in the Sheet menu) **Add Contact** is "Yes", it adds the contact to your Google Contacts.
* Checks for the presence of an input named "mailingList".
    - If it finds such an input and it was checked by the person filling out the form (or has a truthy value) AND the option **Group Email** is not empty, it sends an email to the contact inviting them to join the Google Group.
    - This only works if you already have a Google Group set up. See below for tips on creating a read-only Google Group "mailing list";
* Stores the data in your Google Sheet.
* Sends a notification email to you, the owner of the account in which this script is running.
    -  The reason it sends the notification only to the account owner is that the notification email includes information about the Google Contact entry (if the contact was added to your Google Contacts OR if the Contact already exists). This would be confusing if you send the notification to a different email because the Goolge Contact is in YOUR Contacts, not their's.
    - For information about how to set up email forwarding and forwarding filters in your Gmail account, refer to this documentation:
        * https://support.google.com/a/answer/14724207?hl=en
        * https://support.google.com/mail/answer/6579

## Create a Google Sheet ##

* Sign into your Google account and go to https://docs.google.com/spreadsheets
* Create a new sheet and add the headers: "Timestamp", "First Name", "Last Name", "Email", "Phone", "Mailing List", "Source", "Subject", "Message".

## Create a new Apps Script ##

* Click the menu item [Extensions > Apps Script]
* Delete the default code in the script window that appears.
* Paste the contents of the Code.gs file included in this repo into the script window.

## Name the Project ##

* At the top of the Apps Script project window, click in the box that says "Untitled project" and give it a name.

## Create a "sidebar.html" file ##

* Click the Plus sign (+) beside "Files" at the top of the Apps Script project window and choose "HTML".
* Rename the newly created file "sidebar". Google automatically adds the "html" file extension.
* Paste the content of the "sidebar.html" file included in this repo into the new file.
* Save

## Create a "readme.html" file ##

* Click the Plus sign (+) beside "Files" at the top of the Apps Script project window and choose "HTML".
* Rename the newly created file "readme". Google automatically adds the "html" file extension.
* Paste the content of the "readme.html" file included in this repo into the new file.
* Save

## Add Required Services ##

* Click on the Plus sign (+) beside "Services" on the left side of the screen.
* In the popup that appears, scroll down the list and choose [Peopleapi] to select the People API service.
* Click on the Plus sign again and choose [Groups Settings API].
* Save your project.
* Google may ask for permissions in order to use the script. If this happens, click through the resulting screens.

## Run a Test ##

* At the top of the Apps Script project window, choose "testPost" from the dropdown menu next to "Debug"
* Click on "Run"
* If Google hasn't asked you for permissions to run this script yet, it will now. Click through the screens to give Google the requied permissions.
* Check the email for the account you are currently signed into. You should see a new test email which was sent from the script.

## Add a Trigger ##

* Click on the "Triggers" icon on the left side of the screen (it looks like an alarm clock);
* Click "Add Trigger" and use these settings:
    - Choose which function to run: doPost
    - Choose which deployment should run: Head
    - Select event source: From spreadsheet
    - Select event type: On form submit
* Click "Save"
* Google may ask you to verify more permission, just click through the screens.

## Create a Deployment ##

* If you are still in the "Triggers" screen, go back to the "Editor" screen (The icon looks like [<>]).
* Click on Deploy > New deployment
* Click on the Gear icon next to "Select Type" and choose "Web app".
* For Description, enter something meaningful
* For Execute as, choose "Me"
* For Who has access, choose Anyone
* Click on "Deploy"
* On the resulting screen, you will see your App URL.
* Copy the App URL and save it. You will use it later.

## Store Your APP URL ##

* Near the top of your Code.gs file, is a variable named "appUrl";
* Paste your App URL as the value of this variable. Remember to include quotes around it.
* Save

## Set Options ##

* Go back to the browser window/tab displaying your Google Sheet and refresh the page.
* You should see a new menu item at the top named "ContactSheet" (it may take a few seconds for it to show up).
* Choose ContactSheet > Options.
* A side bar will open, allowing you to choose a few options. If you are unsure what an options means, click on the "?" icon to the right of it.
* Click "Save"

## Create a Google Group mailing list (optional) ##

If you want the script to automatically send the person who submitted the form an invitation to join your Google Group (mailing list) if they check a box in your form named "mailingList", you must have an existing Google Group. See below for tips on how to create a Google Groups mailing list.

## Create an HTML Form ##

* Create a web page and insert an HTML form like the SAMPLE HTML FORM below.
* Enter the App URL as the value for the "action" attribute of your HTML form. The App URL is available from the sidebar that opens when you choose the menu option ContactSheet > Options.
* The form's METHOD must be "POST" ie. form action="[App URL]" method="post"

**NOTE:** The html form input names are CASE SENSITIVE. They must match the same casing as shown below.

Your form can have inputs named: firstName, lastName, email, phone, mailingList, subject, message, mirage. As far as the script is concerned, it only requires firstName and email to function, so if you just need a simple "Join our mailing list" form, you can safely omit the other fields. However, the values for any fields that are absent in the form will display as "undefined" in the sheet.

  | input name | input type  | Required |
  |------------|-------------|----------|
  |firstName   |text         | REQUIRED |
  |lastName    |text         | OPTIONAL |
  |email       |email        | REQUIRED |
  |phone       |tel          | OPTIONAL |
  |mailingList |checkbox     | OPTIONAL |
  |subject     |text         | OPTIONAL |
  |message     |textarea     | OPTIONAL |
  |mirage      |hidden       | OPTIONAL |


## Permissions Required ##

This Apps Script requires the following Google permissions:

* **`https://www.googleapis.com/auth/contacts`**:  To view and manage your contacts (required for adding contacts).
* **`https://www.googleapis.com/auth/groups`**: To check for the existence of the Google Group defined by the parameter groupEmail or the UI option "Group Email" (required for mailing list functionality).
* **`https://www.googleapis.com/auth/script.external_request`**: To communicate with your web server (required for handling form submissions).
* **`https://www.googleapis.com/auth/spreadsheets`**:  To read and write data to your Google Spreadsheets (required for storing form data).
* **`https://www.googleapis.com/auth/userinfo.email`**: To get the email address of the Google account running this script (used for sending notifications and Google Group invitations).

## Tips for creating a Google Group mailing list ##

The key points here are to set the group so **only group managers can post**, **all members can read posts**, and **anyone on the web can auto-join**.

* Sign into your Google account
* Go to groups.google.com
* Click "Create Group"
* Give the group a name, write down its email address, and press "Next"
* On the next screen (Choose privacy settings), set:
    - Who can search for group: Anyone on the web
    - Who can join group: Anyone can join
    - Who can view conversations: Group members (or "Anyone on the web" if you want anyone to be able to view your announcements)
    - Who can post: Group managers **IMPORTANT**
    - Who can view members: Group managers
* On the next screen (Add members), you can assign members, managers and owners if you want.
    - Make sure that "Subscription" is "Each email" so an email is sent out to all members every time you post an announcement
    - Click on "Create group"


## SAMPLE HTML FORM ##

**Note:** The form includes an input named "mirage" which acts as a crude bot trap.

Many simple bots fill and submit forms without executing javascript. In such cases, the "mirage" input will be sent to the server.

The onsubmit trigger attached to the sample form fires a javascript function to remove the field before submitting, so when a person submits the form in a real web browser, the "mirage" input will NOT be sent to the server.

The server checks for the existance of the "mirage" field and rejects the request if it is present.

This will not work if the browser has javascript disabled.

If you don't want to include the mirage input, you can omit it (remember to remove the "onsubmit" trigger from the form as well).

        <noscript>
          <p>In order for this form to behave properly, javascript must be enabled.</p>
        </noscript>

        <form
          method="post"
          onsubmit="event.target.mirage.remove()"
          action="[APP_URL]">

          <div>
            <input
              required
              type="text"
              name="firstName"
              placeholder="First Name"
              aria-label="first name">
          </div>

          <div>
            <input
              type="text"
              name="lastName"
              placeholder="Last Name"
              aria-label="last name">
          </div>

          <div>
            <input
              required
              type="email"
              name="email"
              placeholder="Email"
              aria-label="email">
          </div>

          <div>
            <input
              type="tel"
              name="phone"
              placeholder="Phone"
              aria-label="phone">
          </div>

          <div>
            <input
              required
              type="text"
              name="subject"
              placeholder="Subject"
              aria-label="subject">
          </div>

          <div>
            <textarea
              required
              name="message"
              placeholder="Message"
              aria-label="message">
            </textarea>
          </div>

          <div>
            <input
              type="checkbox"
              name="mailingList"
              id="add-to-list">

            <label for="add-to-list">
              Subscribe to our Mailing List
            </label>
          </div>

          <button type="submit">Submit</button>

          <input required type="hidden" name="mirage" value="bot trap">
        </form>
