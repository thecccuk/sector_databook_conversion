import pandas as pd
import numpy as np

# only import thefuzz if it's installed
try:
    from thefuzz import process
except ModuleNotFoundError:
    process = None

# modelling years
START_YEAR = 2021
END_YEAR = 2050
YEARS = list(range(START_YEAR, END_YEAR+1))

# some constants
SECTOR = 'Industry'
DEVOLVED_AUTHS = ['United Kingdom', 'Scotland', 'Wales', 'Northern Ireland']
GASES = ['CARBON', 'CH4', 'N2O']
SD_COLUMNS = ['Measure ID', 'Country', 'Sector', 'Subsector', 'Measure Name', 'Measure Variable', 'Variable Unit']
CATEGORIES = ['Dispersed or Cluster Site', 'Process', 'Selected Option']
SCENARIO = 'Balanced pathway'
REEE_SHEET = "REEE Projection - EE Sector"

# compute the discount factor for each year as 1/(1+r)^y
SOCIAL_DISCOUNT_RATE = 0.035  # 3.5%
DISCOUNT_FACTORS = {y: 1/(1+SOCIAL_DISCOUNT_RATE)**y for y in YEARS}

# Update SD_COLUMNS to include dynamic categories based on the defined CATEGORIES list.
SD_COLUMNS = SD_COLUMNS[:4] + [f'Category{i+2}: {c}' for i, c in enumerate(CATEGORIES, 1)] + SD_COLUMNS[4:]
SD_COLUMNS += YEARS

def get_subsectors(sector_map: pd.DataFrame, ccc_sector: str):
    """
    Retrieve EE subsectors and their mapping to a given CCC sector from a sector mapping DataFrame.

    Parameters
    ----------
    sector_map : pd.DataFrame
        DataFrame containing sector mappings.
    ccc_sector : str
        The CCC sector for which subsectors are to be retrieved.

    Returns
    -------
    subsectors : list
        List of EE subsectors for the specified CCC sector.
    ee_sector_to_subsector : dict
        Dictionary mapping EE sectors to CCC subsectors.

    Raises
    ------
    ValueError
        If the specified CCC sector is not present in the sector map.
    """
    # Validate the CCC sector against the sector map.
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
    Load and preprocess NZIP output data for a specified CCC sector.

    Parameters
    ----------
    nzip_path : str
        Path to the NZIP output workbook.
    sector_map_path : str
        Path to the sector definitions CSV file.
    sector : str
        The CCC sector for which data is to be loaded.

    Returns
    -------
    pd.DataFrame
        The processed DataFrame containing NZIP data filtered and adjusted for the specified sector.
    """
    sector_map_df = pd.read_csv(sector_map_path)
    
    # Load the NZIP output, specifying the sheet and columns to be used.
    # Warning: this is a bit slow
    with open(nzip_path, 'rb') as f:
        df = pd.read_excel(f, sheet_name='CCC Outputs', header=10, usecols='F:CWV')

    # Ensure the EE sectors from the map and NZIP output match, printing discrepancies.
    ee_sectors_from_map = set(sector_map_df['EE Sector'])
    ee_sectors_from_nzip = set(df['Element_sector'])
    if ee_sectors_from_map != ee_sectors_from_nzip:
        in_map_not_nzip = ee_sectors_from_map - ee_sectors_from_nzip
        in_nzip_not_map = ee_sectors_from_nzip - ee_sectors_from_map
        print(f'EE sectors in map but not NZIP: {in_map_not_nzip}, EE sectors in NZIP but not map: {in_nzip_not_map}')

    # Filter the DataFrame based on the relevant subsectors for the given CCC sector.
    subsectors, subsector_map = get_subsectors(sector_map_df, sector)
    df = df.loc[df['Element_sector'].isin(subsectors)]
    df['CCC Subsector'] = df['Element_sector'].map(subsector_map)

    # Data cleaning steps for string and numeric columns, including handling of NaN values and specific subsector exclusions.
    df['Selected Option'] = df['Selected Option'].fillna('')
    df['Technology Type'] = df['Technology Type'].fillna('')

    fix_numeric_cols = ['% CARBON Emissions', '% CH4 Emissions', '% N2O Emissions']
    fix_numeric_cols += [f'Total AM costs (£m) {y}' for y in YEARS]
    fix_numeric_cols += [f'AM opex (£m) {y}' for y in YEARS]
    fix_numeric_cols += [f'AM fuel costs (£m) {y}' for y in YEARS]
    fix_numeric_cols += [f'Cost Differential (£m) {y}' for y in YEARS]
    fix_numeric_cols += [f'Total direct emissions abated (MtCO2e) {y}' for y in YEARS]
    fix_numeric_cols += [f'Total indirect emissions abated (MtCO2e) {y}' for y in YEARS]
    for col in fix_numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.fillna(0)
    df = df.loc[df['CCC Subsector'] != 'Non-road mobile machinery']

    return df.copy()

def add_cols(df):
    """
    Perform calculations and add additional columns to the DataFrame.

    This is used for the following variables:
    - Additional capital expenditure
    - Additional operational expenditure
    - Additional demand final non bio
    - Low carbon costs
    - Abatement cost new unit
    - Abatement cost average measure

    Parameters
    ----------
    df : pd.DataFrame
        The DataFrame to be modified.

    Returns
    -------
    pd.DataFrame
        The modified DataFrame with additional calculated columns.
    """
    # Calculations for capex, opex, low carbon costs, additional demand, and abatement costs.
    # These include differences between actual and counterfactual expenditures, and adjustments based on implementation year.
    for y in YEARS:
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
        # TODO: what is the cost differential?
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
    Search for a string within the column names of a DataFrame using fuzzy matching.

    Parameters
    ----------
    df : pandas.DataFrame
        The DataFrame whose columns are to be searched.
    search_string : str
        The string to search for within the column names.
    limit : int, optional
        The maximum number of matches to return.

    Returns
    -------
    list
        A list of tuples containing the matched column names and their corresponding scores, limited by the specified count.
    """
    # Use fuzzy matching to find the best matches for the search_string in the DataFrame's column names.
    return process.extract(search_string, df.columns.astype(str), limit=limit)

def sector_databook_format(df, variable_name, variable_unit):
    """
    Format the DataFrame according to the sector databook specifications.

    Parameters
    ----------
    df : pandas.DataFrame
        The DataFrame to be formatted.
    variable_name : str
        The name of the variable for the "Measure Variable" column.
    variable_unit : str
        The unit of measurement for the "Variable Unit" column.

    Returns
    -------
    pd.DataFrame
        The formatted DataFrame with specific columns adjusted to match sector databook requirements.
    """
    df = df.reset_index()
    df['Measure ID'] = ''
    df['Sector'] = SECTOR
    df['Subsector'] = df['CCC Subsector']
    df['Measure Name'] = df['Measure Technology']
    df['Measure Variable'] = variable_name
    df['Variable Unit'] = variable_unit
    for i, category in enumerate(CATEGORIES):
        df[f'Category{i+3}: {category}'] = df[category]
    df = df[SD_COLUMNS]
    return df

def aggregate_timeseries_country(df, timeseries, variable_name, variable_unit, weight_col=None, country='United Kingdom', scale=None, measure=None):
    """
    Aggregate timeseries data for a specific country and variable, optionally applying weighting and scaling.

    Parameters
    ----------
    df : pandas.DataFrame
        The DataFrame containing timeseries data.
    timeseries : str
        The name of the timeseries variable.
    variable_name : str
        The name of the variable to be used in the aggregation.
    variable_unit : str
        The unit for the aggregated variable.
    weight_col : str, optional
        The column name to use for weighting the data. Default is None.
    country : str, optional
        The country for which data is aggregated. Default is 'United Kingdom'.
    scale : float, optional
        A scaling factor to apply to the data. Default is None.
    measure : str, optional
        A specific measure to filter the data by. Default is None.

    Returns
    -------
    pd.DataFrame
        The aggregated DataFrame with timeseries data summed up per specified parameters.
    """
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
    if not measure:
        tech_map = {'Blue Hydrogen': 'Hydrogen', 'Green Hydrogen': 'Hydrogen', 'Electric': 'Electrification'}
        df['Measure Technology'] = df['Technology Type'].replace(tech_map)
    else:
        df['Measure Technology'] = measure
        
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
    """
    Aggregate timeseries data for all countries specified in the DEVOLVED_AUTHS list.

    Parameters
    ----------
    df : pandas.DataFrame
        The DataFrame containing timeseries data.
    **kwargs
        Additional keyword arguments to be passed to `aggregate_timeseries_country`.

    Returns
    -------
    pd.DataFrame
        A DataFrame containing the aggregated timeseries data for all specified countries.
    """
    # Aggregate data for each country in DEVOLVED_AUTHS and combine the results.
    dfs = [aggregate_timeseries_country(df, country=country, **kwargs) for country in DEVOLVED_AUTHS]
    df = pd.concat(dfs)
    return df

def add_reee(nzip_path, df, baseline_col, post_reee_col, out_col, usecols="E:AL", header=327, nrows=28):
    df = df.copy()
    # read the energy efficiency data from the nzip model.
    # here we taking the "People" scenario which matches the balanced pathway
    with open(nzip_path, 'rb') as f:
        index = pd.read_excel(f, sheet_name=REEE_SHEET, header=header, nrows=nrows, usecols='D', index_col=0)
        ee_df = pd.read_excel(f, sheet_name=REEE_SHEET, header=header, nrows=nrows, usecols=usecols, index_col=None)
        ee_df.index = index.index
    ee_df.columns = [int(x[:4]) for x in ee_df.columns] # cast column names to int and fix names
    ee_df = ee_df[YEARS] # select only relevant years
    
    # for some reason the ee fraction is report in % for the abatement, but in factors of 1.x for the demands
    # so we need to account for this here
    #if baseline_col != "Baseline emissions (MtCO2e)":
    #    ee_df = 1 - ee_df # 22/03/24: removing this for now as it seems to be incorrect, instead use the EE fracs from emissions

    for y in YEARS:
        # ee_frac represents the percentage reduction in emissions due to EE
        ee_frac = df['Element_sector'].map(ee_df[y])
        ee = (df[f'{post_reee_col} {y}'] / (1 - ee_frac)) - df[f'{post_reee_col} {y}']
        re = (df[f'{baseline_col} {y}'] - df[f'{post_reee_col} {y}']) - ee        

        # when computing demands, we flip the sign as these are "additional demands" rather than "abated emissions"
        if baseline_col != "Baseline emissions (MtCO2e)":
            ee = -ee
            re = -re

        df[f'RE {out_col} {y}'] = re
        df[f'EE {out_col} {y}'] = ee

        # assert neither of these are nan or inf
        assert np.isfinite(df[f'RE {out_col} {y}']).all()
        assert np.isfinite(df[f'EE {out_col} {y}']).all()

    return df


def sd_measure_level(df, args, reee_args=None, baseline=True, nzip_path=None):
    """
    Process and aggregate measure-level data according to specified arguments.

    Parameters
    ----------
    df : pandas.DataFrame
        The DataFrame containing the data to be processed.
    args : list
        A list of dictionaries specifying the arguments for data aggregation.
    baseline : bool, optional
        If True, exclude measures not unimplemented measures.

    Returns
    -------
    pd.DataFrame
        The aggregated and processed DataFrame with measure-level data.
    """
    if not baseline:
        df = df.loc[df['Year of Implementation'] < 2060].copy()

    sd_df = pd.DataFrame(columns=SD_COLUMNS)
    for kwargs in args:
        if sd_df.empty:
            sd_df = aggregate_timeseries(df, **kwargs)
        else:
            sd_df = pd.concat([sd_df, aggregate_timeseries(df, **kwargs)])

    # handle REEE measures
    if not baseline:
        if reee_args is not None:
            reee_args = reee_args.copy()
        for kwargs in reee_args:
            agg_kwargs = {'variable_name': kwargs['out_col'], 'variable_unit': kwargs.pop('variable_unit'), 'scale': kwargs.pop('scale', None)}
            reee_df = add_reee(nzip_path, df, **kwargs)
            sd_df = pd.concat([sd_df, aggregate_timeseries(reee_df, timeseries=f"RE {kwargs['out_col']}", measure='Resource Efficiency', **agg_kwargs)])
            sd_df = pd.concat([sd_df, aggregate_timeseries(reee_df, timeseries=f"EE {kwargs['out_col']}", measure='Energy Efficiency', **agg_kwargs)])

    #
    # compute some additional costs:
    #

    # compute "Abatement cost new unit" as:
    # the "cost differential" in a given year divided by "total emissions abated" in that year
    cost = sd_df['Measure Variable'] == f'cost differential'
    abated = sd_df['Measure Variable'] == f'total emissions abated'
    sd_df.loc[cost, YEARS] = (sd_df.loc[cost, YEARS] / sd_df.loc[abated, YEARS]).fillna(0)
    sd_df.loc[cost, 'Measure Variable'] = f'Abatement cost new unit'
    sd_df = sd_df.loc[~abated] # remove intermediate rows used in the calculation

    # compute "Abatement cost average measure" as:
    # the cumulative "cost differential" divided by the cumulative "total emissions abated"
    cost = sd_df['Measure Variable'] == f'cum cost differential'
    abated = sd_df['Measure Variable'] == f'cum total emissions abated'
    sd_df.loc[cost, YEARS] = (sd_df.loc[cost, YEARS] / sd_df.loc[abated, YEARS]).fillna(0)
    sd_df.loc[cost, 'Measure Variable'] = f'Abatement cost average measure'
    sd_df = sd_df.loc[~abated] # remove intermediate rows used in the calculation

    #assert not sd_df.duplicated().any()
    return sd_df

def baseline_from_measure_level(df):
    """
    Convert measure-level data to baseline formatting.

    Parameters
    ----------
    df : pandas.DataFrame
        The DataFrame containing measure-level data.

    Returns
    -------
    pd.DataFrame
        A DataFrame formatted to baseline specifications, aggregating data as necessary.
    """
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
    Aggregate additional demand for a specified fuel and technology.

    Parameters
    ----------
    df : pandas.DataFrame
        The DataFrame containing raw NZIP data.
    agg_df : pd.DataFrame
        The aggregate DataFrame to which the results will be added.
    fuel : str
        The name of the fuel for which additional demand is calculated.
    fuel_out : str, optional
        The output column name for the aggregated demand. Defaults to `fuel`.
    tech : str, optional
        The technology type to filter on. Default is 'CCS'.

    Returns
    -------
    pd.DataFrame
        The updated aggregate DataFrame with additional demand for the specified fuel.
    """
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
    """
    Create an aggregate DataFrame containing results from various calculations.

    Parameters
    ----------
    df : pandas.DataFrame
        The raw DataFrame to be aggregated.
    measure_level_kwargs : list
        Keyword arguments for processing measure-level data.
    baseline_kwargs : list
        Keyword arguments for processing baseline data.
    sector : str
        The sector for which the aggregation is performed.

    Returns
    -------
    pd.DataFrame
        An aggregate DataFrame with summarized data across specified measures and baselines.
    """
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
