## msal-20-spfx

## Summary
This webpart demosntrates the authentication and access token acquisition with [MSAL.js 2.0](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-browser) inside SharePoint Framework (SPFx). On a button click it tries three login options
* silent
* popup
* redirect
and in case a login already took place previously it first tries to acquire in access token from cache or per refresh.

## msal-20-spfx in action
![WebPartInAction](https://mmsharepoint.files.wordpress.com/2020/08/04mailresult.png)

A detailed functionality and technical description can be found in the [author's blog post](https://mmsharepoint.wordpress.com/2020/08/15/using-msal-js-2-0-in-sharepoint-framework-spfx/)

## Used SharePoint Framework Version

![drop](https://img.shields.io/badge/drop-1.11.0-green.svg)

## Applies to
Usage of [MSAL.js 2.0 Authorization Code Flow](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-browser)

## Solution

Solution|Author(s)
--------|---------
outlook-2-sp-spfx| Markus Moeller ([@moeller2_0](http://www.twitter.com/moeller2_0))

## Version history

Version|Date|Comments
-------|----|--------
1.0|August 15, 2020|Initial release
||

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome
* Clone this repository
* in the command line run:
  * restore dependencies: `npm install`
  * build solution `gulp build --ship`
  * bundle solution: `gulp bundle --ship`
  * package solution: `gulp package-solution --ship`
  * locate solution at `.\sharepoint\solution\msal-20-spfx.sppkg` 
  * upload it to your tenant app catalog
  * Register an app in Azure AD as SPA with a redirect URI
  * Install your webpart on a given site 
  * Instantiate your webpart on a page in that site
  * Configure it with app id, your redirect URI (best: That page's url) and your tenant domain (YOURTENANT.onmicrosoft.com)

## Features

This webpart shows the following capabilities on top of the SharePoint Framework:

* MSAL 2.0 authorization code flow including
  * silent login
  * popup login
  * redirect login
  * token acquisition