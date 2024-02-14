import pandas as pd
from thefuzz import process

# modelling years
START_YEAR = 2021
END_YEAR = 2050
YEARS = list(range(START_YEAR, END_YEAR+1))

# some constants
DEVOLVED_AUTHS = ['United Kingdom', 'Scotland', 'Wales', 'Northern Ireland']
GASES = ['CARBON', 'CH4', 'N2O']
SD_COLUMNS = ['Measure ID', 'Country', 'Sector', 'Subsector', 'Measure Name', 'Measure Variable', 'Variable Unit']
CATEGORIES = ['Dispersed or Cluster Site', 'Process', 'Selected Option']#, 'Traded / non-traded']
SCENARIO = 'Balanced pathway'

# compute the discount factor for each year as 1/(1+r)^y
SOCIAL_DISCOUNT_RATE = 0.035 # 3.5%
DISCOUNT_FACTORS = {y: 1/(1+SOCIAL_DISCOUNT_RATE)**y for y in YEARS}

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
    fix_numeric_cols += [f'Cost Differential (£m) {y}' for y in YEARS]
    fix_numeric_cols += [f'Total direct emissions abated (MtCO2e) {y}' for y in YEARS]
    fix_numeric_cols += [f'Total indirect emissions abated (MtCO2e) {y}' for y in YEARS]
    for col in fix_numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # any NaN (not a number) values are set to 0
    df = df.fillna(0)

    # manually drop NRMM subsector
    df = df.loc[df['CCC Subsector'] != 'Non-road mobile machinery']

    return df.copy()

def add_cols(df):
    """
    Do some intermediate calculations and add some columns to the dataframe.

    This is used for the following variables:
    - Additional capital expenditure
    - Additional operational expenditure
    - Additional demand final non bio
    - Low carbon costs
    - Abatement cost new unit
    - Abatement cost average measure
    """
    # add costs
    for y in YEARS:
        # additional capex and opex are calculated as the difference between the AM and counterfactual costs
        df[f'capex {y}'] = df[f'AM capex (£m) {y}'] - df[f'Counterfactual capex (£m) {y}']
        df[f'opex {y}'] = df[f'AM opex (£m) {y}'] + df[f'AM fuel costs (£m) {y}'] - (df[f'Counterfactual opex (£m) {y}'] + df[f'Counterfactual fuel costs (£m) {y}'])

        # low carbon costs are calculated as follows:
        # 1. if the year of implementation is less than the year in question, the costs are 0
        # 2. otherwise, the costs are the same as the total AM capex and opex columns
        df[f'capex low carbon {y}'] = df[f'AM capex (£m) {y}'].copy()
        df.loc[df['Year of Implementation'] < y, f'capex low carbon {y}'] = 0
        df[f'opex low carbon {y}'] = df[f'AM opex (£m) {y}'].copy()
        df.loc[df['Year of Implementation'] < y, f'opex low carbon {y}'] = 0
    
        # additional demand final non bio, calculated as follows:
        # 1. based on the process/sector, we know the fraction of non bio waste
        # 2. we multiply this by the total solid fuel use to get the total non bio waste
        # 3. we then subtract the post REEE total non bio waste to get the change in non bio waste before and after REEE
        non_bio_waste_dict = {'Kiln - Cement': 0.23, 'Kiln - Lime': 0.23, 'Incinerators': 1.0, 'Other Chemicals': 0.54}
        frac_bio_waste = df['Process'].copy().map(non_bio_waste_dict).fillna(0)
        frac_bio_waste.loc[df['Element_sector'] == 'Other Chemicals'] = non_bio_waste_dict['Other Chemicals']
        df[f'total non bio waste {y}'] = df[f'Total solid fuel use (GWh) {y}'] * frac_bio_waste
        df[f'post REEE total non bio waste {y}'] = df[f'Post REEE baseline in solid fuel use (GWh) {y}'] * frac_bio_waste
        df[f'Change in non bio waste {y}'] = df[f'total non bio waste {y}'] - df[f'post REEE total non bio waste {y}']
        
        # Abatement cost new unit: cost differential in each year divided by total emissions abated in each year
        abatement = df[f'Total direct emissions abated (MtCO2e) {y}'] + df[f'Total indirect emissions abated (MtCO2e) {y}']
        cost = df[f'Cost Differential (£m) {y}']
        df[f'total emissions abated {y}'] = abatement
        df[f'cost differential {y}'] = cost

        # Abatement cost average measure: cumulative cost differential divided by cumulative total emissions abated
        if y == START_YEAR:
            df[f'cum cost differential {y}'] = cost
            df[f'cum total emissions abated {y}'] = abatement
        else:
            df[f'cum cost differential {y}'] = df[f'cum cost differential {y-1}'] + cost
            df[f'cum total emissions abated {y}'] = df[f'cum total emissions abated {y-1}'] + abatement

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

def sd_measure_level(df, args_list, drop_not_implemented=False):
    
    # for non-baseline data, we can drop rows that have an implementation year >2060
    if drop_not_implemented:
        df = df.loc[df['Year of Implementation'] < 2060].copy()

    # process measure level data
    sd_df = pd.DataFrame(columns=SD_COLUMNS)
    for kwargs in args_list:
        if sd_df.empty:
            sd_df = aggregate_timeseries(df, **kwargs)
        else:
            sd_df = pd.concat([sd_df, aggregate_timeseries(df, **kwargs)])

    # get some extra costs
    cost = sd_df['Measure Variable'] == f'cost differential'
    abated = sd_df['Measure Variable'] == f'total emissions abated'

    # divide numeric columns
    sd_df.loc[cost, YEARS] = (sd_df.loc[cost, YEARS] / sd_df.loc[abated, YEARS]).fillna(0)
    sd_df.loc[cost, 'Measure Variable'] = f'Abatement cost new unit'
    sd_df = sd_df.loc[~abated]

    # now do the same for the cumulative columns
    cost = sd_df['Measure Variable'] == f'cum cost differential'
    abated = sd_df['Measure Variable'] == f'cum total emissions abated'
    sd_df.loc[cost, YEARS] = (sd_df.loc[cost, YEARS] / sd_df.loc[abated, YEARS]).fillna(0)
    sd_df.loc[cost, 'Measure Variable'] = f'Abatement cost average measure'
    sd_df = sd_df.loc[~abated]

    assert not sd_df.duplicated().any()
    return sd_df

def baseline_from_measure_level(df):
    """Reformat a measure level table to match the baseline formatting."""
    cols = ['Country', 'Sector', 'Subsector', 'Measure Name', 'Measure Variable', 'Variable Unit']
    cols = cols[:4] + [f'Category{i+2}: {c}' for i, c in enumerate(CATEGORIES, 1)] + cols[4:]
    bl_df = df.groupby(cols).sum(numeric_only=True)
    bl_df = bl_df.reset_index()
    bl_df = bl_df.rename(columns={'Measure Variable': 'Baseline Variable'})
    bl_df = bl_df.drop(columns=['Measure Name'])
    assert not bl_df.duplicated().any()
    return bl_df

def get_additional_demand_agg(df, agg_df, fuel, fuel_out=None, tech='CCS'):
    """
    Get the aggregate additional demand for a given fuel.

    Parameters
    ----------
    df : pd.DataFrame
        The raw NZIP dataframe.
    agg_df : pd.DataFrame
        The aggregate dataframe to add the results to.
    fuel : str
        The name of the fuel to get the additional demand for (nzip column name).
    fuel_out : str
        The name of the fuel to use in the aggregate dataframe (SD column name).
        If not provided, assume this is the same as fuel.
    tech : str
        The technology type to filter on.
    """
    # if fuel_out is not provided, assume it is the same as fuel
    if fuel_out is None:
        fuel_out = fuel

    # only consider rows for the given technology type
    ccs_df = df.loc[df['Technology Type'] == tech].copy()
    year_of_implementation = ccs_df['Year of Implementation']

    # for each year
    for y in YEARS:
        # compute the total fuel use after the year of implementation, and multiply by the abatement rate
        ccs_df[f'{fuel} use after implementation {y}'] = ccs_df[f'Total {fuel} use (GWh) {y}'].copy()
        ccs_df.loc[y < year_of_implementation, f'{fuel} use after implementation {y}'] = 0
        
        # updated guidance from CB team: don't multiply by the CCS capture rate
        #ccs_df[f'{fuel} use after implementation {y}'] *= ccs_df['Abatement Rate']
        
        # sum all rows and convert from GWh to TWh
        agg_df.loc[f'Additional demand {fuel_out} abated', y] = ccs_df[f'{fuel} use after implementation {y}'].sum() * 0.001
    
    agg_df.loc[f'Additional demand {fuel_out} abated', 'Aggregate Variable'] = f'Additional demand {fuel_out} abated'
    agg_df.loc[f'Additional demand {fuel_out} abated', 'Variable Unit'] = 'TWh'
    agg_df.loc[f'Additional demand {fuel_out} abated', 'Scenario'] = SCENARIO

    return agg_df

def get_aggregate_df(df, measure_level_kwargs, baseline_kwargs, sector):
    df = df.copy()

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
    agg_df.loc['Direct emissions total', 'Scenario'] = SCENARIO

    # traded
    agg_df.loc['Baseline traded emissions total'] = total_baseline_emissions_traded
    agg_df.loc['Baseline traded emissions total', 'Aggregate Variable'] = 'Baseline traded emissions total'
    agg_df.loc['Baseline traded emissions total', 'Variable Unit'] = 'MtCO2e'
    agg_df.loc['Baseline traded emissions total', 'Scenario'] = 'Baseline'

    agg_df.loc['Direct traded emissions total'] = total_pathway_emissions_traded
    agg_df.loc['Direct traded emissions total', 'Aggregate Variable'] = 'Direct traded emissions total'
    agg_df.loc['Direct traded emissions total', 'Variable Unit'] = 'MtCO2e'
    agg_df.loc['Direct traded emissions total', 'Scenario'] = SCENARIO

    # additional demand gas, petroleum, solid fuel
    agg_df = get_additional_demand_agg(df, agg_df, 'natural gas', 'gas')
    agg_df = get_additional_demand_agg(df, agg_df, 'petroleum')
    agg_df = get_additional_demand_agg(df, agg_df, 'solid fuel')
 
    # fill some missing stuff
    agg_df['Country'] = 'United Kingdom'
    agg_df['Sector'] = sector

    return agg_df
