import streamlit as st
import io
import warnings
warnings.simplefilter(action='ignore')
import pandas as pd
from datetime import datetime,timezone
import numpy as np
# from natsort import index_natsorted #for version sorting of hierarchical numbers


st.title("KNS BOM Parser")

if 'bom_df' not in st.session_state:
    st.session_state.bom_df = None

if 'bom_file' not in st.session_state:
    st.session_state.bom_file = None

if 'output_bom_df' not in st.session_state:
    st.session_state.output_bom_df = None

if 'output_bom_file' not in st.session_state:
    st.session_state.output_bom_file = io.StringIO()

def parse_oracle_bom(bom_file_obj):
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", category=UserWarning)
        bom_df = pd.read_excel(bom_file_obj,sheet_name=0,engine="openpyxl",skiprows=0,usecols='A:U',converters={
            'BOM_LEVEL':str,
            'ITEM':str,
            'MANUFACTURING_ITEM':str,
            'QTY':float,
            'POS':str,
        })
    bom_df = bom_df.dropna(subset=['BOM_LEVEL'])
    bom_df['System No.'] = np.nan
    bom_df['Drawing Reference'] = ''
    bom_df['Unit Cost [SGD]'] = np.nan
    bom_df['Total Cost [SGD]'] = np.nan
    bom_df['WIP or Released'] = 'WIP'
    bom_df.loc[bom_df['BOM_LEVEL']=='TOP MODEL : ', 'Hierarchical No.'] = '1'
    bom_df.loc[bom_df['BOM_LEVEL']=='TOP MODEL : ', 'Obsolete'] = 'N'
    bom_df.loc[bom_df['BOM_LEVEL']=='TOP MODEL : ', 'QTY'] = 1.0
    bom_df.loc[bom_df['BOM_LEVEL']=='TOP MODEL : ', 'UOM'] = 'PCS'
    bom_df.loc[bom_df['BOM_LEVEL']=='TOP MODEL : ', 'MANUFACTURER_NAME'] = 'AKRIBIS ASSY'
    bom_df.loc[bom_df['BOM_LEVEL']=='TOP MODEL : ', 'Parent'] = 'None'

    # Populate hierarchical number
    for i in bom_df.index:
        bom_df = populate_hier_num(bom_df,i)

    # Determine parent
    bom_df['Parent'] = bom_df['Hierarchical No.'].apply(lambda s: '.'.join(s.split('.')[:-1]))
    bom_df.loc[bom_df['BOM_LEVEL']=='TOP MODEL : ', 'Parent'] = 'None'
    bom_df.loc[bom_df['BOM_LEVEL']=='TOP MODEL : ', 'System No.'] = 'CMMKNS'
    bom_df.loc[bom_df['BOM_LEVEL']=='TOP MODEL : ', 'Description\n(Order Part No / Dwg No / REV No.)'] = bom_df.loc[bom_df['BOM_LEVEL']=='TOP MODEL : ', 'ITEM']

    # Formulate part number (description 1)
    bom_df.loc[bom_df['MANUFACTURER_NAME'].isna(),'Description\n(Order Part No / Dwg No / REV No.)'] = bom_df.loc[bom_df['MANUFACTURER_NAME'].isna(),'ITEM'] + 'REV' + bom_df.loc[bom_df['MANUFACTURER_NAME'].isna(),'REV']
    bom_df.loc[~bom_df['MANUFACTURER_NAME'].isna(),'Description\n(Order Part No / Dwg No / REV No.)'] = bom_df.loc[~bom_df['MANUFACTURER_NAME'].isna(),'MANUFACTURER_PART_NUMBER']

    # Assign best-guess manufacturer
    bom_df.loc[~bom_df['MANUFACTURER_NAME'].isna(),'Manufacturer'] = bom_df.loc[~bom_df['MANUFACTURER_NAME'].isna(),'MANUFACTURER_NAME']
    bom_df.loc[bom_df['MANUFACTURER_NAME'].isna() & bom_df['ITEM_DESCRIPTION'].str.startswith('ASSY') ,'Manufacturer'] = 'AKRIBIS ASSY'
    bom_df.loc[bom_df['MANUFACTURER_NAME'].isna() & bom_df['ITEM_DESCRIPTION'].str.contains('CABLE COMPLEMENT') ,'Manufacturer'] = 'AKRIBIS ASSY'
    bom_df.loc[bom_df['MANUFACTURER_NAME'].isna() & bom_df['ITEM_DESCRIPTION'].str.startswith('CBL_') ,'Manufacturer'] = 'AKRIBIS CABLING'
    bom_df.loc[bom_df['MANUFACTURER_NAME'].isna() & bom_df['ITEM_DESCRIPTION'].str.startswith('TERM_') ,'Manufacturer'] = 'AKRIBIS CABLING'
    bom_df.loc[bom_df['MANUFACTURER_NAME'].isna() & bom_df['Hierarchical No.'].apply(lambda s: not (s in set(bom_df['Parent']))),'Manufacturer'] = 'AKRIBIS FAB'

    # Assign system number categories
    bom_df.loc[bom_df['System No.'].isna() & bom_df['ITEM_DESCRIPTION'].str.startswith('ASSY') ,'System No.'] = 'SASSSM'
    bom_df.loc[bom_df['System No.'].isna() & bom_df['ITEM_DESCRIPTION'].str.contains('CABLE COMPLEMENT') ,'System No.'] = 'SASSSM'
    bom_df.loc[bom_df['System No.'].isna() & bom_df['ITEM_DESCRIPTION'].str.startswith('CBL_') ,'System No.'] = 'AACACW'
    bom_df.loc[bom_df['System No.'].isna() & bom_df['ITEM_DESCRIPTION'].str.startswith('TERM_') ,'System No.'] = 'AACACW'
    bom_df.loc[bom_df['System No.'].isna() & (bom_df['Manufacturer'] == 'AKRIBIS FAB'),'System No.'] = 'FAB'
    bom_df.loc[bom_df['System No.'].isna() & (bom_df['ITEM_DESCRIPTION'].str.startswith('WIRE')),'System No.'] = 'EEPCNW'
    bom_df.loc[bom_df['System No.'].isna() & (bom_df['ITEM_DESCRIPTION'].str.startswith('SCREW')),'System No.'] = 'MEPFSC'
    bom_df.loc[bom_df['System No.'].isna() & (bom_df['ITEM_DESCRIPTION'].str.startswith('WASHER')),'System No.'] = 'MEPFSC'
    
    bom_df['Description 2\n(Description / Dwg Title)'] = bom_df['ITEM_DESCRIPTION']
    bom_df['Qty'] = bom_df['QTY']

    bom_df = bom_df[['Hierarchical No.','System No.','Description\n(Order Part No / Dwg No / REV No.)','Description 2\n(Description / Dwg Title)','Qty','UOM','Unit Cost [SGD]','Total Cost [SGD]','Manufacturer','Drawing Reference','WIP or Released','Obsolete','Parent']]
    return bom_df

def populate_hier_num(bom_df,i):
    cur_bom_df = bom_df.copy(deep=True)
    # print('Index %d' % i)
    if not pd.isna(cur_bom_df.at[i,'Hierarchical No.']): #do nothing if hier num already exists
        # print('skipped')
        return cur_bom_df
    parent_hier_num_list = cur_bom_df.loc[cur_bom_df['ITEM']==cur_bom_df.at[i,'MANUFACTURING_ITEM'],'Hierarchical No.']
    # print(parent_hier_num_list)
    if len(parent_hier_num_list) > 1:
        raise ValueError('Duplicate parent found')
    if len(parent_hier_num_list) < 1:
        raise ValueError('Parent not found.')
    parent_hier_num = parent_hier_num_list.iloc[0]
    if pd.isna(parent_hier_num): #recursively assign hierarchical number for the parent if it does not yet exist
        # print('recursion')
        cur_bom_df = populate_hier_num(cur_bom_df,bom_df.index(cur_bom_df['ITEM'] == cur_bom_df.iloc[i]['MANUFACTURING_ITEM']).iloc[0])
        parent_hier_num = cur_bom_df.loc[cur_bom_df['ITEM']==cur_bom_df.iloc[i]['MANUFACTURING_ITEM'],'Hierarchical No.'].iloc[0]
    siblings_hier_nums = cur_bom_df.loc[cur_bom_df['Hierarchical No.'].str.startswith(parent_hier_num+'.') &
                                        (~cur_bom_df['Hierarchical No.'].isna()) &
                                        (cur_bom_df['BOM_LEVEL']==cur_bom_df.at[i,'BOM_LEVEL']) &
                                        (~(cur_bom_df['Obsolete']=='Y'))
                                        ,'Hierarchical No.']
    siblings_hier_nums = siblings_hier_nums.apply(lambda s: int(s.split(parent_hier_num+'.')[1]))
    if len(siblings_hier_nums) == 0:
        current_item_hier_num = parent_hier_num + '.1'
        cur_bom_df.at[i,'Obsolete'] = 'N'
    else:
        duplicate_siblings_positions = cur_bom_df.loc[cur_bom_df['Hierarchical No.'].str.startswith(parent_hier_num+'.') &
                                            (~cur_bom_df['Hierarchical No.'].isna()) &
                                            (cur_bom_df['BOM_LEVEL']==cur_bom_df.at[i,'BOM_LEVEL']) &
                                            (~(cur_bom_df['Obsolete']=='Y')) &
                                            ((cur_bom_df['POS']==cur_bom_df.at[i,'POS']))
                                            ,'POS']
        # print(duplicate_siblings_positions)
        if len(duplicate_siblings_positions) > 0:
            cur_bom_df.at[i,'Obsolete'] = 'Y'
            current_item_hier_num = parent_hier_num + '.' + str(siblings_hier_nums.max())
        else:
            cur_bom_df.at[i,'Obsolete'] = 'N'
            current_item_hier_num = parent_hier_num + '.' + str(siblings_hier_nums.max()+1)
    
    # print('Parent Hier Num: %s' % parent_hier_num)
    # print('Siblings Hier Num: %s' % siblings_hier_nums)
    # print('Current Hier Num: %s' % current_item_hier_num)

    #TO-DO: Deal with alternate parts in Format D BOM

    
    cur_bom_df.at[i,'Hierarchical No.'] = current_item_hier_num
    return cur_bom_df

def parse_bom():
    if st.session_state.bom_file is None:
        st.session_state["upload_state"] = "Upload a file first!"
    else:
        st.session_state.bom_df = parse_oracle_bom(st.session_state.bom_file)

def output_bom():
    if st.session_state.bom_df is None:
        st.session_state["upload_state"] = "Upload a file first!"
    else:
        with pd.ExcelWriter(st.session_state.output_bom_file,engine='xlsxwriter') as writer:
            workbook = writer.book
            header_format = workbook.add_format(
            {
                "bold": True,
                "text_wrap": True,
                "valign": "top",
                "border": 1,
            }
        )
            st.session_state.bom_df.to_excel(writer, sheet_name='System BOM')
            (max_row, max_col) = st.session_state.bom_df.shape
            writer.sheets['System BOM'].autofilter(0, 0, max_row, max_col)

st.session_state.bom_file = st.file_uploader('Drop source KNS BOM here.', type='xlsx', accept_multiple_files=False, key='source_bom_upload', label_visibility="visible")
st.button("Parse BOM", on_click=parse_bom)
# st.button("Output BOM", on_click=output_bom)
# st.download_button('Download BOM', st.session_state.output_bom_file, file_name='output_bom.txt')
st.dataframe(data=st.session_state.bom_df)
