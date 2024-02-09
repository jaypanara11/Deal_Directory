import { WebPartContext } from "@microsoft/sp-webpart-base";
import { graph } from "@pnp/graph";
import { sp, PeoplePickerEntity, ClientPeoplePickerQueryParameters, SearchQuery, SearchResults, SearchProperty, SortDirection } from '@pnp/sp';
import { PrincipalType } from "@pnp/sp/src/sitegroups";
import { isRelativeUrl } from "office-ui-fabric-react";
import { ISPServices } from "./ISPServices";
import { ITickerProperties } from "./ITickerProperties";
import { Web } from "sp-pnp-js";

var siteUrl = "https://hpsinvestmentpartnersllc.sharepoint.com/sites/Deals";
var rootUrl_or = window.location.origin;

if (rootUrl_or.toLowerCase() == "https://realitytechhub.sharepoint.com") {
    siteUrl = "https://realitytechhub.sharepoint.com/sites/HPS";
}

export class spservices implements ISPServices {
    constructor(private context: WebPartContext) {
        sp.setup({
            spfxContext: this.context
        });
    }

    public async searchTickers(searchString: string): Promise<ITickerProperties[]> {
        let tickers: any[];
        searchString = searchString == "All" ? '' : searchString;

        var rootUrl = window.location.origin;
        var industryColumnID = "h9dd1fe92117459f8478f2ff5d4fda81";
        var countryColumnID = "pa5dbdc3da8a47d998513bdf7c8e29f4";
        if (rootUrl.toLowerCase() == "https://realitytechhub.sharepoint.com") {
            industryColumnID = "d662fec64c7f4dff8c0144757cd53e85";
            countryColumnID = "dfa5fe335876420fb13ff5385d7c3f12";
        }


        try {
            if (searchString) {
                let web: any = new Web(siteUrl);
                tickers = await web.lists.getByTitle("Deals Central List").items.select("CrossPlatformStatus", "ProjectName_x002d_Link", "ProjectName", "DealTeamRegion", "Fund", "Issuer_x002d_Link", "Issuer", "Platform", "Title", "Salesforce_x002d_Link", industryColumnID, countryColumnID).filter("substringof('" + searchString + "',Issuer)").getAll();
            }
            else {
                let web: any = new Web(siteUrl);
                tickers = await web.lists.getByTitle("Deals Central List").items.select("CrossPlatformStatus", "ProjectName_x002d_Link", "ProjectName", "DealTeamRegion", "Fund", "Issuer", "Issuer_x002d_Link", "Platform", "Title", "Salesforce_x002d_Link", industryColumnID, countryColumnID).getAll();
            }

            let tickerArr: ITickerProperties[] = tickers.map((value, index) => {
                var itemUrl = "#", industry = "", country = "", sfUrl = "";
                if (value.Issuer_x002d_Link != null) {
                    //itemUrl = value.URL.Url;
                    itemUrl = value.Issuer_x002d_Link.Url;
                }
                if (value.Salesforce_x002d_Link != null) {
                    sfUrl = value.Salesforce_x002d_Link.Url;
                }
                if (value[industryColumnID] != "" && value[industryColumnID] != null) {
                    var indVal = value[industryColumnID];
                    industry = indVal.split('|')[0];
                }
                if (value[countryColumnID] != "" && value[countryColumnID] != null) {
                    var cntryVal = value[countryColumnID];
                    country = cntryVal.split('|')[0];
                }

                return {
                    Title: value.Issuer,
                    Sector: value.Platform,
                    URL: itemUrl,
                    Issuer: value.Issuer,
                    ProjectName: value.ProjectName,
                    Fund: value.Fund,
                    CrossPlatformStatus: value.CrossPlatformStatus,
                    DealTeamRegion: value.DealTeamRegion,
                    Industry: industry,
                    Country: country,
                    SalesforceLink: value.Salesforce_x002d_Link,
                    DisplayType: ""
                }
            })

            return tickerArr;
        } catch (error) {
            Promise.reject(error);
        }
    }

    public async _getImageBase64(pictureUrl: string): Promise<string> {
        return new Promise((resolve, reject) => {
            let image = new Image();
            image.addEventListener("load", () => {
                let tempCanvas = document.createElement("canvas");
                tempCanvas.width = image.width,
                    tempCanvas.height = image.height,
                    tempCanvas.getContext("2d").drawImage(image, 0, 0);
                let base64Str;
                try {
                    base64Str = tempCanvas.toDataURL("image/png");
                } catch (e) {
                    return "";
                }
                resolve(base64Str);
            });
            image.src = pictureUrl;
        });
    }

    public async searchTickersNew(DisplayType: string, searchString: string, srchQry: string, isInitialSearch: boolean, sectorFilter?: string, issuerFilter?: string, projectnameFilter?: string, fundFilter?: string, crossPlatformStatusFilter?: string, dealTeamRegionFilter?: string, industryFilter?: string, countryFilter?: string): Promise<ITickerProperties[]> {
        let tickers: any[];
        searchString = searchString == "All" ? '' : searchString;
        srchQry = srchQry == "All" ? '' : srchQry;
        var searchField = "Issuer";
        if (DisplayType == "Deal") {
            searchField = "Title";
        }

        var rootUrl = window.location.origin;
        var industryColumnID = "h9dd1fe92117459f8478f2ff5d4fda81";
        var countryColumnID = "pa5dbdc3da8a47d998513bdf7c8e29f4";
        if (rootUrl.toLowerCase() == "https://realitytechhub.sharepoint.com") {
            industryColumnID = "d662fec64c7f4dff8c0144757cd53e85";
            countryColumnID = "dfa5fe335876420fb13ff5385d7c3f12";
        }


        try {
            if (!isInitialSearch && (searchString || srchQry || sectorFilter || issuerFilter || projectnameFilter || fundFilter || crossPlatformStatusFilter || dealTeamRegionFilter || industryFilter || countryFilter)) {
                let searchText = searchString != "" ? searchString : srchQry;
                // 07/02/24 start
                searchText = encodeURIComponent(searchText);
                // 07/02/24 end
                if (searchText) {
                    var SectorFilterString = sectorFilter ? (" and Platform eq '" + sectorFilter + "'") : "";
                    var issuerFilterString = issuerFilter ? (" and Issuer eq '" + issuerFilter + "'") : "";
                    //var issuerFilterString = "substringof('" + issuerFilter + "',Issuer)";
                    if (searchText != "" || issuerFilter == undefined) {
                        issuerFilterString = "";
                    }
                    //var prjnameFilterString = projectnameFilter ? (" and ProjectName eq '" + projectnameFilter + "'") : "";
                    var prjnameFilterString = "substringof('" + projectnameFilter + "',ProjectName)";
                    if (projectnameFilter == undefined || projectnameFilter == "") {
                        prjnameFilterString = "";
                    }

                    var fundFilterString = fundFilter ? (" and Fund eq '" + fundFilter + "'") : "";
                    var CrossPlatformFilterString = crossPlatformStatusFilter ? (" and CrossPlatformStatus eq '" + crossPlatformStatusFilter + "'") : "";
                    var dealTeamRegnFilterString = dealTeamRegionFilter ? (" and DealTeamRegion eq '" + dealTeamRegionFilter + "'") : "";
                    var industryFilterString = industryFilter ? (" and TaxCatchAll/Term eq '" + industryFilter + "'") : "";
                    var countryFilterString = countryFilter ? (" and TaxCatchAll/Term eq '" + countryFilter + "'") : "";


                    if (searchString) {
                        let web: any = new Web(siteUrl);
                        tickers = await web.lists.getByTitle("Deals Central List").items.select("CrossPlatformStatus", "ProjectName", "ProjectName_x002d_Link", "DealTeamRegion", "Fund", "Issuer", "Platform", "Title", "Issuer_x002d_Link", "Salesforce_x002d_Link", industryColumnID, countryColumnID).filter("substringof('" + searchText + "'," + searchField + ")" + SectorFilterString + issuerFilterString + prjnameFilterString + fundFilterString + CrossPlatformFilterString + dealTeamRegnFilterString + countryFilterString).orderBy("Title", true).getAll();
                    }
                    else {
                        // tickers = await sp.web.lists.getByTitle("Deals Central List").items.filter("startswith(Title,'" + searchText + "')" + SectorFilterString + "").orderBy("Title", true).get();
                        let web: any = new Web(siteUrl);
                        tickers = await web.lists.getByTitle("Deals Central List").items.select("CrossPlatformStatus", "ProjectName", "DealTeamRegion", "Fund", "Issuer", "Platform", "Issuer_x002d_Link", "Title", "Salesforce_x002d_Link", "ProjectName_x002d_Link", industryColumnID, countryColumnID).filter("substringof('" + searchText + "'," + searchField + ")" + SectorFilterString + issuerFilterString + prjnameFilterString + fundFilterString + CrossPlatformFilterString + dealTeamRegnFilterString + industryFilterString + countryFilterString).orderBy("Title", true).getAll();
                    }

                }
                else {
                    // tickers = await sp.web.lists.getByTitle("Deals Central List").items.select("CrossPlatformStatus", "ProjectName", "DealTeamRegion", "Fund", "Issuer", "Platform", "Title", "ProjectName_x002d_Link", "dfa5fe335876420fb13ff5385d7c3f12", "d662fec64c7f4dff8c0144757cd53e85").filter("Platform eq '" + sectorFilter + "' or Issuer eq '" + issuerFilter + "' or ProjectName eq '" + projectnameFilter + "' or Fund eq '" + fundFilter + "' or CrossPlatformStatus eq '" + crossPlatformStatusFilter + "' or DealTeamRegion eq '" + dealTeamRegionFilter + "' or TaxCatchAll/Term eq '" + industryFilter + "' or TaxCatchAll/Term eq '" + countryFilter + "'").orderBy("Title", true).get();
                    // tickers = await sp.web.lists.getByTitle("Deals Central List").items.select("CrossPlatformStatus", "ProjectName", "DealTeamRegion", "Fund", "Issuer", "Platform", "Issuer_x002d_Link", "Title", "Salesforce_x002d_Link", "ProjectName_x002d_Link", "dfa5fe335876420fb13ff5385d7c3f12", "d662fec64c7f4dff8c0144757cd53e85").filter("Platform eq '" + sectorFilter + "' or substringof('" + issuerFilter + "',Issuer) or substringof('" + projectnameFilter + "',ProjectName) or Fund eq '" + fundFilter + "' or CrossPlatformStatus eq '" + crossPlatformStatusFilter + "' or DealTeamRegion eq '" + dealTeamRegionFilter + "' or TaxCatchAll/Term eq '" + industryFilter + "' or TaxCatchAll/Term eq '" + countryFilter + "'").orderBy("Title", true).get();
                    var filterString = "";
                    if (sectorFilter != undefined && sectorFilter != "") {
                        filterString = "Platform eq '" + sectorFilter + "'";
                    }
                    if (issuerFilter != undefined && issuerFilter != "") {
                        if (filterString == "") {
                            filterString = "Issuer eq '" + issuerFilter + "'";
                        }
                        else {
                            filterString = filterString + " and Issuer eq '" + issuerFilter + "'";
                        }
                    }
                    if (projectnameFilter != undefined && projectnameFilter != "") {
                        if (filterString == "") {
                            filterString = "ProjectName eq '" + projectnameFilter + "'";
                        }
                        else {
                            filterString = filterString + " and ProjectName eq '" + projectnameFilter + "'";
                        }
                    }
                    if (fundFilter != undefined && fundFilter != "") {
                        if (filterString == "") {
                            filterString = "Fund eq '" + fundFilter + "'";
                        }
                        else {
                            filterString = filterString + " and Fund eq '" + fundFilter + "'";
                        }
                    }
                    if (crossPlatformStatusFilter != undefined && crossPlatformStatusFilter != "") {
                        if (filterString == "") {
                            filterString = "CrossPlatformStatus eq '" + crossPlatformStatusFilter + "'";
                        }
                        else {
                            filterString = filterString + " and CrossPlatformStatus eq '" + crossPlatformStatusFilter + "'";
                        }
                    }
                    if (dealTeamRegionFilter != undefined && dealTeamRegionFilter != "") {
                        if (filterString == "") {
                            filterString = "DealTeamRegion eq '" + dealTeamRegionFilter + "'";
                        }
                        else {
                            filterString = filterString + " and DealTeamRegion eq '" + dealTeamRegionFilter + "'";
                        }
                    }
                    if (industryFilter != undefined && industryFilter != "") {
                        if (filterString == "") {
                            filterString = "TaxCatchAll/Term eq '" + industryFilter + "'";
                        }
                        else {
                            filterString = filterString + " and TaxCatchAll/Term eq '" + industryFilter + "'";
                        }
                    }
                    if (countryFilter != undefined && countryFilter != "") {
                        if (filterString == "") {
                            filterString = "TaxCatchAll/Term eq '" + countryFilter + "'";
                        }
                        else {
                            filterString = filterString + " and TaxCatchAll/Term eq '" + countryFilter + "'";
                        }
                    }

                    let web: any = new Web(siteUrl);
                    tickers = await web.lists.getByTitle("Deals Central List").items.select("CrossPlatformStatus", "ProjectName", "DealTeamRegion", "Fund", "Issuer", "Platform", "Issuer_x002d_Link", "Title", "Salesforce_x002d_Link", "ProjectName_x002d_Link", industryColumnID, countryColumnID).filter(filterString).orderBy("Title", true).getAll();
                }
            }
            else {
                if (searchString) {
                    //tickers = await sp.web.lists.getByTitle("Deals Central List").items.filter("startswith(Title,'" + searchString + "')").orderBy("Title", true).get();
                    let web: any = new Web(siteUrl);
                    tickers = await web.lists.getByTitle("Deals Central List").items.select("CrossPlatformStatus", "ProjectName", "DealTeamRegion", "Fund", "Issuer", "Platform", "Issuer_x002d_Link", "Title", "Salesforce_x002d_Link", "ProjectName_x002d_Link", industryColumnID, countryColumnID).filter("startswith(" + searchField + ",'" + searchString + "')").orderBy("Title", true).getAll();
                }
                else {
                    let web: any = new Web(siteUrl);
                    tickers = await web.lists.getByTitle("Deals Central List").items.select("CrossPlatformStatus", "ProjectName", "DealTeamRegion", "Fund", "Issuer", "Platform", "Issuer_x002d_Link", "Title", "Salesforce_x002d_Link", "ProjectName_x002d_Link", industryColumnID, countryColumnID).orderBy("Title", true).getAll();
                    //.select("Title", "Description")
                    //dfa5fe335876420fb13ff5385d7c3f12  = country
                    //d662fec64c7f4dff8c0144757cd53e85 = industry

                    //live industry - h9dd1fe92117459f8478f2ff5d4fda81
                    //live country - pa5dbdc3da8a47d998513bdf7c8e29f4
                }
            }

            var tempTicker = [];
            if (DisplayType == "Deal") {
                tempTicker = tickers;
            }
            else {
                if (tickers.length > 0) {
                    for (var i = 0; i < tickers.length; i++) {
                        var isAvail = false;
                        for (var j = 0; j < tempTicker.length; j++) {
                            if (tempTicker[j].Issuer == tickers[i].Issuer) {
                                isAvail = true;
                            }
                        }
                        if (isAvail == false) {
                            tempTicker.push(tickers[i]);
                        }
                    }
                }
            }



            let tickerArr: ITickerProperties[] = tempTicker.map((value, index) => {
                var itemUrl = "#", industry = "", country = "", sfUrl = "";
                if (DisplayType == "Deal") {
                    if (value.ProjectName_x002d_Link != null) {
                        itemUrl = value.ProjectName_x002d_Link.Url;
                    }
                }
                else {
                    if (value.Issuer_x002d_Link != null) {
                        itemUrl = value.Issuer_x002d_Link.Url;
                    }
                }

                if (value.Salesforce_x002d_Link != null) {
                    sfUrl = value.Salesforce_x002d_Link.Url;
                }
                if (value[industryColumnID] != "" && value[industryColumnID] != null) {
                    var indVal = value[industryColumnID];
                    industry = indVal.split('|')[0];
                }
                if (value[countryColumnID] != "" && value[countryColumnID] != null) {
                    var cntryVal = value[countryColumnID];
                    country = cntryVal.split('|')[0];
                }
                return {
                    Title: value.Title,
                    Sector: value.Platform,
                    URL: itemUrl,
                    Issuer: value.Issuer,
                    ProjectName: value.ProjectName,
                    Fund: value.Fund,
                    CrossPlatformStatus: value.CrossPlatformStatus,
                    DealTeamRegion: value.DealTeamRegion,
                    Industry: industry,
                    Country: country,
                    SalesforceLink: sfUrl,
                    DisplayType: DisplayType
                }
            })

            return tickerArr;
        } catch (error) {
            Promise.reject(error);
        }
    }
}
