## dash

This is a SharePoint list data visualization web part built in SharePoint Framework (SPFx) using the React SPFx template. This web part utilizes [React](https://reactjs.org/), [SharePoint REST web services](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service), and [Chart.js](http://www.chartjs.org/).

### Building Your Own Web Part

This solution is intended to accompany [Introduction to SharePoint Framework](https://sharepointfx.io/), an online educational course that helps you to learn modern SharePoint Framework development techniques. Learn how to build your own dashboard web part by following the lessons found at [sharepointfx.io](https://sharepointfx.io/).

### Getting Started

```bash
# Install dependencies
npm i

# Run the local workbench
gulp serve
```

### Deploying to SharePoint

```bash
# Bundle the solution
gulp bundle --ship

# Package the solution
#  - This creates a sharepoint/solution/dash.sppkg file
gulp package-solution --ship
```

Once you have a `dash.sppkg` file, you can deploy this to your SharePoint environment's [App Catalog](https://docs.microsoft.com/en-us/sharepoint/use-app-catalog). See the **Deploying and Updating Solutions** lesson for more information on solution deployment.

### Learn More
For more information about the structure and functionality of this solution, see the [official SharePoint Framework documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview).
