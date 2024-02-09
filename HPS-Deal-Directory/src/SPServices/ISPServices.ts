import { PeoplePickerEntity } from '@pnp/sp';

export interface ISPServices {

    searchTickers(searchString: string);
    searchTickersNew(DisplayType: string, searchString: string, srchQry: string, isInitialSearch: boolean, sectorFilter?: string, issuerFilter?: string, projectnameFilter?: string, fundFilter?: string, crossPlatformFilter?: string, dealTeamRegionFilter?: string, industryFilter?: string, countryFilter?: string);

}