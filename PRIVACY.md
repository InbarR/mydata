# Privacy

MyData is designed to process data locally.

- Files selected through the browser remain in browser memory.
- Workbook contents are not uploaded to a server.
- MyData has no backend, analytics, telemetry, advertising, or remote logging.
- Temporary browser `blob:` URLs are used only within the current page.
- Refreshing or closing the page clears imported data from the app.

The public website hosts only MyData's static application code.

## Exports

Excel, CSV, and HTML exports contain data from the current view. You are
responsible for storing and sharing exports according to the sensitivity and
licensing of the source data.

## Sensitive files

Before processing confidential, regulated, or sensitivity-labeled files,
confirm that using a public web origin complies with your organization's
policies. If policy prohibits it, use the offline `mydata.html` artifact in an
approved environment.
