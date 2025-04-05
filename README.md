# Bulk Upload SPFx

A simple SPFx solution built with React for bulk data addition to SharePoint lists. Originally developed for adding testing data quickly, this solution can also be useful in projects that require bulk data addition.

## Overview

This solution leverages:

- **React** for the user interface.
- **PnPjs** for data operations with SharePoint.
- **SheetJS** to read data from Excel files.

It provides two upload methods:

1. **Sequential Add:** Uploads one item at a time using async/await.
2. **Parallel Add:** Uploads all items concurrently using `Promise.all`.
   To make it more reusable for you guys, I customized it a little bit and added these two uploading methods.

## Requirements

- [Node.js](https://nodejs.org/) (v18.x.x or above)
- [PnPjs](https://pnp.github.io/pnpjs/)
- [SheetJS](https://sheetjs.com/)
