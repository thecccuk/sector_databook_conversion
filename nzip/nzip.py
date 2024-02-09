import pandas as pd
from thefuzz import process


START_YEAR = 2021
END_YEAR = 2050
YEARS = list(range(START_YEAR, END_YEAR+1))
DEVOLVED_AUTHS = ['United Kingdom', 'Scotland', 'Wales', 'Northern Ireland']
GASES = ['CARBON', 'CH4', 'N2O']
SD_COLUMNS = ['Measure ID', 'Country', 'Sector', 'Subsector', 'Measure Name', 'Measure Variable', 'Variable Unit']
CATEGORIES = ['Dispersed or Cluster Site', 'Process']#, 'Traded / non-traded']

SD_COLUMNS = SD_COLUMNS[:4] + [f'Category{i+2}: {c}' for i, c in enumerate(CATEGORIES, 1)] + SD_COLUMNS[4:]
SD_COLUMNS += YEARS


def get_subsectors(sector_map: pd.DataFrame, ccc_sector: str):
    """
    Get the subsectors (EE sectors) relevant for a given CCC sector.

    Parameters
    ----------
    sector_map : pd.DataFrame
        The sector mapping to use.
    ccc_sector : str
        The CCC sector to get the subsectors for.
    """
    # check the ccc sector is valid
    if ccc_sector not in sector_map['CCC Sector'].unique():
        raise ValueError(f'Invalid CCC Sector: {ccc_sector}, must be one of {sector_map["CCC Sector"].unique()}')

    # filter rows based on the CCC Sector we are interested in
    sector_defs = sector_map.loc[sector_map['CCC Sector'] == ccc_sector]

    # get the EE sectors which correspond to this CCC sector
    subsectors = sector_defs['EE Sector'].unique().tolist()

    # get the mapping from ee sector to ccc subsector
    ee_sector_to_subsector = sector_defs.set_index('EE Sector')['CCC Subsector'].to_dict()

    return subsectors, ee_sector_to_subsector

def load_nzip(nzip_path: str, sector_map_path: str, sector: str):
    """
    Load NZIP outputs and filter rows based on CCC sector.
    Some initial data cleaning is also applied.

    Parameters
    ----------
    nzip_path : str
        The path to the NZIP output workbook.
    sector_map_path : str
        The path to the NZIP sector definitions csv file.
    sector : str
        CCC sector to load. Must be one of 'Industry', 'Fuel Supply', or 'Waste'.
    """

    # get the mapping from ee sector to ccc subsector
    sector_map_df = pd.read_csv(sector_map_path)
    
    # read the NZIP output workbook. this can take a minute or two, we could speed it up by first converting it to a csv
    with open(nzip_path, 'rb') as f:
        df = pd.read_excel(f, sheet_name='CCC Outputs', header=10, usecols='F:CWV')

    # check that the ee sectors are consistent
    ee_sectors_from_map = set(sector_map_df['EE Sector'])
    ee_sectors_from_nzip = set(df['Element_sector'])
    if ee_sectors_from_map != ee_sectors_from_nzip:
        in_map_not_nzip = ee_sectors_from_map - ee_sectors_from_nzip
        in_nzip_not_map = ee_sectors_from_nzip - ee_sectors_from_map
        print(f'EE sectors in map but not NZIP: {in_map_not_nzip}, EE sectors in NZIP but not map: {in_nzip_not_map}')

    # get the subsectors relevant for the given sector
    subsectors, subsector_map = get_subsectors(sector_map_df, sector)

    # select relevant rows based on the sector
    df = df.loc[df['Element_sector'].isin(subsectors)]

    # add a column for the CCC subsector
    df['CCC Subsector'] = df['Element_sector'].map(subsector_map)

    # fix string columns which have some empty cells
    df['Selected Option'] = df['Selected Option'].fillna('')
    df['Technology Type'] = df['Technology Type'].fillna('')

    # fix some numeric columns which have some non-numeric values (these values will be set to 0 later)
    fix_numeric_cols = ['% CARBON Emissions', '% CH4 Emissions', '% N2O Emissions']
    fix_numeric_cols += [f'Total AM costs (£m) {y}' for y in YEARS]
    fix_numeric_cols += [f'AM opex (£m) {y}' for y in YEARS]
    fix_numeric_cols += [f'AM fuel costs (£m) {y}' for y in YEARS]
    for col in fix_numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # any NaN (not a number) values are set to 0
    df = df.fillna(0)

    # add capex and opex cols
    for y in range(START_YEAR, END_YEAR+1):
        df[f'capex {y}'] = df[f'AM capex (£m) {y}'] - df[f'Counterfactual capex (£m) {y}']
        df[f'opex {y}'] = df[f'AM opex (£m) {y}'] + df[f'AM fuel costs (£m) {y}'] - (df[f'Counterfactual opex (£m) {y}'] + df[f'Counterfactual fuel costs (£m) {y}'])

    # manually drop NRMM subsector
    df = df.loc[df['CCC Subsector'] != 'Non-road mobile machinery']

    return df

def col_search(df, search_string, limit=5):
    """
    Search the columns of a dataframe for a string using thefuzz.

    Parameters
    ----------
    df : pandas.DataFrame
        The dataframe to search.
    search_string : str
        The string to search for.
    limit : int
        The maximum number of results to return.
    """
    return process.extract(search_string, df.columns.astype(str), limit=limit)

def sector_databook_format(df, variable_name, variable_unit):
    df = df.reset_index()
    df['Measure ID'] = ''
    df['Sector'] = 'Industry'
    df['Subsector'] = df['CCC Subsector']
    df['Measure Name'] = df['Measure Technology']
    df['Measure Variable'] = variable_name
    df['Variable Unit'] = variable_unit
    for i, category in enumerate(CATEGORIES):
        df[f'Category{i+3}: {category}'] = df[category]
    df = df[SD_COLUMNS]
    return df

def aggregate_timeseries_country(df, timeseries, variable_name, variable_unit, weight_col=None, country='United Kingdom', scale=None):

    if country != 'United Kingdom':
        # filter to rows for the given country
        df = df.loc[df['Country'] == country].copy()

    # get the emissions time series columns
    total_emissions_cols = [f'{timeseries} {y}' for y in YEARS]
    emissions_cols = YEARS
    df[emissions_cols] = df[total_emissions_cols].copy()

    # multiply by another column and/or then scale by a fixed value
    if weight_col:
        df[emissions_cols] = df[emissions_cols].multiply(df[weight_col], axis=0)
    if scale:
        df[emissions_cols] = df[emissions_cols] * scale

    # map some technology types to a new name
    tech_map = {'Blue Hydrogen': 'Hydrogen', 'Green Hydrogen': 'Hydrogen', 'Electric': 'Electrification'}
    df['Measure Technology'] = df['Technology Type'].replace(tech_map)
        
    # sum rows corresponding to the same measure
    agg_emissions_df = df.groupby(['CCC Subsector', 'Measure Technology'] + CATEGORIES)[emissions_cols].sum()

    # add country column
    agg_emissions_df['Country'] = country

    # format as sector databook
    df = sector_databook_format(agg_emissions_df, variable_name, variable_unit)

    # drop rows where each year is 0
    #df = df.loc[(df[YEARS] != 0).any(axis=1)]

    return df

def aggregate_timeseries(df, **kwargs):
    # go through each country and combine the results
    dfs = [aggregate_timeseries_country(df, country=country, **kwargs) for country in DEVOLVED_AUTHS]
    df = pd.concat(dfs)
    return df

def sd_measure_level(df, args_list):
    sd_df = pd.DataFrame(columns=SD_COLUMNS)
    for kwargs in args_list:
        if sd_df.empty:
            sd_df = aggregate_timeseries(df, **kwargs)
        else:
            sd_df = pd.concat([sd_df, aggregate_timeseries(df, **kwargs)])
    assert not sd_df.duplicated().any()
    return sd_df

def baseline_from_measure_level(df):
    cols = ['Country', 'Sector', 'Subsector', 'Measure Name', 'Measure Variable', 'Variable Unit']
    cols = cols[:4] + [f'Category{i+2}: {c}' for i, c in enumerate(CATEGORIES, 1)] + cols[4:]
    bl_df = df.groupby(cols).sum(numeric_only=True)
    bl_df = bl_df.reset_index()
    bl_df = bl_df.rename(columns={'Measure Variable': 'Baseline Variable'})
    bl_df = bl_df.drop(columns=['Measure Name'])
    assert not bl_df.duplicated().any()
    return bl_df

def get_aggregate_df(df, measure_level_kwargs, baseline_kwargs, sector):
    # create a dataframe to store the aggregate results
    agg_df = pd.DataFrame(columns=['Scenario', 'Country', 'Sector', 'Aggregate Variable', 'Variable Unit'] + list(range(START_YEAR, END_YEAR+1)))

    # get total emissions
    sd_df = sd_measure_level(df, measure_level_kwargs)
    bl_df = sd_measure_level(df, baseline_kwargs)
    bl_df = baseline_from_measure_level(bl_df)    
    
    # get traded emissions
    df_traded = df.loc[df['Traded / non-traded'] == 'traded'].copy()
    sd_df_traded = sd_measure_level(df_traded, measure_level_kwargs)
    bl_df_traded = sd_measure_level(df_traded, baseline_kwargs)
    bl_df_traded = baseline_from_measure_level(bl_df_traded)

    # compute pathway emissions for the UK
    total_abatement = sd_df.loc[(sd_df['Measure Variable'] == 'Abatement total direct') & (sd_df['Country'] == 'United Kingdom')].sum(numeric_only=True)
    total_baseline_emissions = bl_df.loc[(bl_df['Baseline Variable'] == 'Baseline emissions CO2') & (bl_df['Country'] == 'United Kingdom')].sum(numeric_only=True)
    total_pathway_emissions = total_baseline_emissions - total_abatement

    # same for traded
    total_abatement_traded = sd_df_traded.loc[(sd_df_traded['Measure Variable'] == 'Abatement total direct') & (sd_df_traded['Country'] == 'United Kingdom')].sum(numeric_only=True)
    total_baseline_emissions_traded = bl_df_traded.loc[(bl_df_traded['Baseline Variable'] == 'Baseline emissions CO2') & (bl_df_traded['Country'] == 'United Kingdom')].sum(numeric_only=True)
    total_pathway_emissions_traded = total_baseline_emissions_traded - total_abatement_traded
    
    # fill cells manually
    agg_df.loc['Baseline emissions total'] = total_baseline_emissions
    agg_df.loc['Baseline emissions total', 'Aggregate Variable'] = 'Baseline emissions total'
    agg_df.loc['Baseline emissions total', 'Variable Unit'] = 'MtCO2e'
    agg_df.loc['Baseline emissions total', 'Scenario'] = 'Baseline'

    agg_df.loc['Direct emissions total'] = total_pathway_emissions
    agg_df.loc['Direct emissions total', 'Aggregate Variable'] = 'Direct emissions total'
    agg_df.loc['Direct emissions total', 'Variable Unit'] = 'MtCO2e'
    agg_df.loc['Direct emissions total', 'Scenario'] = 'Balanced pathway'

    # traded
    agg_df.loc['Baseline traded emissions total'] = total_baseline_emissions_traded
    agg_df.loc['Baseline traded emissions total', 'Aggregate Variable'] = 'Baseline traded emissions total'
    agg_df.loc['Baseline traded emissions total', 'Variable Unit'] = 'MtCO2e'
    agg_df.loc['Baseline traded emissions total', 'Scenario'] = 'Baseline'

    agg_df.loc['Direct traded emissions total'] = total_pathway_emissions_traded
    agg_df.loc['Direct traded emissions total', 'Aggregate Variable'] = 'Direct traded emissions total'
    agg_df.loc['Direct traded emissions total', 'Variable Unit'] = 'MtCO2e'
    agg_df.loc['Direct traded emissions total', 'Scenario'] = 'Balanced pathway'

    # fill some missing stuff
    agg_df['Country'] = 'United Kingdom'
    agg_df['Sector'] = sector

    return agg_df
