# ExcelToHtml
Converter that converts Microsoft Excel Workbooks into HTML webpages.

## Project Modularization

### ExcelToHtmlConverter

The UI for the Application (WPF)
Model-View-ViewModel Concept

### ExcelToHtmlConverter.Api

API that defines the interface for the business logic

### ExcelToHtmlConverter.Core

The actual business logic of the converter.
Contains the parser to parse Excel-Workbooks into an intermediate model.
Renders the intermediate model with the a Razor-Template to a HTML file.

## Build

```
msbuild -t:Clean,Build
```
