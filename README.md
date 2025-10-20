🏆 District 91 Automated Certificate Generator

This project is a Google Apps Script solution designed to automatically generate certificates for District 91 Toastmasters events using Google Forms and Google Slides.

Whenever a participant submits a Google Form, the script:

1. Retrieves the submission data from Google Sheets.
2. Creates a copy of a Google Slides certificate template.
3. Replaces placeholders (like <<Name>>, <<Event>>, <<Date>>) with the participant’s information.
4. Converts the final certificate into PDF format.
5. Emails the certificate directly to the participant.

🚀 Features

* Fully automated certificate generation from form responses
* Easy to customize with your own Google Slides templates
* Sends certificates automatically via email
* Logs successful submissions and errors
* Ideal for Toastmasters clubs, workshops, or any event with participants

⚙️ How It Works

1) Set up a Google Form to collect participant information.
2) Link the form to a Google Sheet.
3) Create a Google Slides template with placeholders.
4) Copy the provided Apps Script into the linked Sheet.
5) Set up a trigger to run the script on form submission (onFormSubmit).
6) Certificates are generated and emailed automatically!
