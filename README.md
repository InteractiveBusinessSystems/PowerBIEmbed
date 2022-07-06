## power-bi-embed-reports

This is a custom single-page SharePoint web part built with SharePoint Framework and React. This web part was created for Maryville Academy to embed audienced Power BI Reports in a SharePoint page. The solution uses Azure AD registrations for authenticating to Azure AD and an Azure Function for retrieving the embed token and embed urls from Power BI for the reports.

### PreReqs
Node V10.19.0
Gulp V3.9.1

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO
