__author__ = "Mitranshu Kumar"
__email__ = "mitranshu13@gmail.com"
__status__ = "WIP"


# Worksheets
# Windows
# Calculations
# Parameters
# Filters
# Slicers
# Charts

import xml.etree.ElementTree as ET
import html
import sys
import pandas as pd
import string
import random

import xml.etree.ElementTree as ET
import html
import sys
import pandas as pd
import string
import random

def create_ordinal(last_step, df_types, wht = "col", csv_id = None, filename = None):
    ordinal = -1
    for _,i in df_types.iterrows():
        ordinal = ordinal + 1
        col_name = i['index']
        datatype = i[0]

        if datatype == "int64":
            datatype = 'integer'
            remote_type = 20
            aggregation = 'Sum'
        elif datatype == "float64" :
            datatype = 'real'
            remote_type = 5
            aggregation = 'Sum'
        else:
            datatype = 'string'
            remote_type = 129
            aggregation = 'Count'
        
        str_ordinal = str(ordinal)
        if wht== "col":
            out_name =  ET.SubElement(last_step, 'column', {'datatype':datatype, 'name':col_name, 'ordinal':str_ordinal})
        
        if wht == "rm":
            mr_1 = ET.SubElement(last_step, 'metadata-record', {'class':'column'} )  
            rm_1 = ET.SubElement(mr_1, 'remote-name')
            rm_1.text = col_name

            rm_2 = ET.SubElement(mr_1, 'remote-type')
            rm_2.text = str(remote_type)

            rm_3 = ET.SubElement(mr_1, 'local-name')
            rm_3.text = f'[{col_name}]'

            rm_4 = ET.SubElement(mr_1, 'parent-name')
            rm_4.text = f'[{filename}]'

            rm_5 = ET.SubElement(mr_1, 'remote-alias')
            rm_5.text = f'{col_name}'

            rm_6 = ET.SubElement(mr_1, 'ordinal')
            rm_6.text = str(ordinal)

            rm_7 = ET.SubElement(mr_1, 'local-type')
            rm_7.text = datatype

            rm_8 = ET.SubElement(mr_1, 'aggregation')
            rm_8.text = aggregation

            if datatype == 'string':
                rm_8_1 = ET.SubElement(mr_1, 'scale')
                rm_8_1.text = str(1)
                rm_8_2 = ET.SubElement(mr_1, 'width')
                rm_8_2.text = str(1073741823)

            rm_9 = ET.SubElement(mr_1, 'contains-null')
            rm_9.text = 'true'

            if datatype == 'string':
                rm_9_1 = ET.SubElement(mr_1, 'collation', {'flag':'0', 'name':'LEN_RUS'})

            out_name = ET.SubElement(mr_1, '_.fcp.ObjectModelEncapsulateLegacy.true...object-id')
            out_name.text = f'[{csv_id}]'
        
    return out_name

class tableau:
    sheet_cnt = 0
    a1 = 10
    a2 = 1000
    
    def __init__(self,filepath, aliases = None, drill_path = None, parameter = None, calculations = None):
        self.filepath = filepath
        self.aliases = aliases
        self.drill_path = drill_path
        self.parameter = parameter
        self.calculations = calculations
        
    @classmethod
    def update_sht(cls, k):
        cls.sheet_cnt +=1
        cls.a1 +=1
        cls.a2 +=1
        return cls.sheet_cnt, cls.a1 + k, cls.a2 + k
        
    def sheets(self, ws_name, col_for_cal, what_to_cal, color_var = None,
               apply_filter = None, 
               customized_label = True,
               chart_type = "h",
               display_field_labels = True,
               axisline_visibility = False, 
               dropline_visibility = False, refline_visibility = False,
               gridline_col_visibility = False, gridline_row_visibility = False, zeroline_row_visibility = False,
               trendline_visibility = False, bar_color = True, text_label = True, font_family = 'Segoe UI',
               font_size = 10, font_weight = 'regular', text_align_h = 'center', text_align_v = 'center',
               axis_title = False, viewpoint = "entire-view", grid_line = False, color_bars = False
              ):
        sheet_cnt, a1, a2  = tableau.update_sht(1)
        return ws_name, col_for_cal, what_to_cal, sheet_cnt, a1, a2, apply_filter, \
               customized_label, chart_type, display_field_labels, \
               axisline_visibility, dropline_visibility, refline_visibility, \
               gridline_col_visibility, gridline_row_visibility, zeroline_row_visibility, \
               trendline_visibility, bar_color, text_label, font_family, \
               font_size, font_weight, text_align_h, text_align_v, \
               axis_title, viewpoint, grid_line, color_bars, sheet_cnt, a1, a2, color_var
    
    def create_dashboard(self, output_name, shts):
        aliases = self.aliases
        drill_path = self.drill_path
        parameter = self.parameter
        calculations = self.calculations
        
        directory = '/'.join(self.filepath.split('/')[0:-1])
        filename = self.filepath.split('/')[-1]
        filename_name = filename.replace(".","_")
        table_name = filename.replace(".","#")
        caption_txt = filename.replace(".csv","")

        fid = 'federated.0907uli1uxftok1b7adon0cvis26'
        txt_id = 'textscan.1h5akhd1xyggk01bazydt098jrb5'
        csv_id = f'{filename}_6E59FDAB97944C9D9444C65032EEA144'
        
        workbook = ET.Element('workbook', 
                     {
                         'original-version':'18.1',
                         'source-build': '2022.3.4 (20223.23.0214.1014)',
                         'source-platform':'win',
                         'version': '18.1',
                         'xmlns:user': 'http://www.tableausoftware.com/xml/user', 
                     }
                    )

        format_change = ET.SubElement(workbook, "document-format-change-manifest")
        fc = ET.SubElement(format_change, "_.fcp.AccessibleZoneTabOrder.true...AccessibleZoneTabOrder")
        fc = ET.SubElement(format_change, "_.fcp.AnimationOnByDefault.true...AnimationOnByDefault")
        fc = ET.SubElement(format_change, "AutoCreateAndUpdateDSDPhoneLayouts")
        fc = ET.SubElement(format_change, "BasicButtonObject")
        fc = ET.SubElement(format_change, "BasicButtonObjectTextSupport")
        fc = ET.SubElement(format_change, "_.fcp.MarkAnimation.true...MarkAnimation")
        fc = ET.SubElement(format_change, "_.fcp.ObjectModelEncapsulateLegacy.true...ObjectModelEncapsulateLegacy")
        fc = ET.SubElement(format_change, "_.fcp.ObjectModelTableType.true...ObjectModelTableType")
        fc = ET.SubElement(format_change, "_.fcp.SchemaViewerObjectModel.true...SchemaViewerObjectModel")
        fc = ET.SubElement(format_change, "SetMembershipControl")
        fc = ET.SubElement(format_change, "SheetIdentifierTracking")
        fc = ET.SubElement(format_change, "WindowsPersistSimpleIdentifiers")
        fc = ET.SubElement(format_change, "WorksheetBackgroundTransparency")

        preferences = ET.SubElement(workbook, "preferences")
        pn = ET.SubElement(preferences, 'preference', {'name':'ui.discoverpane.show', 'value':'false'})
        pn = ET.SubElement(preferences, 'preference', {'name':'ui.encoding.shelf.height', 'value':'24'})
        pn = ET.SubElement(preferences, 'preference', {'name':'ui.shelf.height', 'value':'26'})

        animation = ET.SubElement(workbook, "_.fcp.AnimationOnByDefault.false...style")
        an = ET.SubElement(animation, '_.fcp.AnimationOnByDefault.false..._.fcp.MarkAnimation.true...style-rule',
                           {'element':'animation'})
        an = ET.SubElement(an, '_.fcp.AnimationOnByDefault.false...format',
                           {'attr':'animation-on', 'value':'ao-on'})

        datasources = ET.SubElement(workbook, 'datasources')
        datasource = ET.SubElement(datasources, 'datasource', 
                                   {'caption': filename_name,'inline': 'true', 'name': f'{fid}', 'version':'18.1'})

        connection = ET.SubElement(datasource,'connection',{'class':'federated'})
        named_connection = ET.SubElement(connection,'named-connections')
        nc = ET.SubElement(named_connection, 'named-connection', {'caption':filename_name, 'name': f'{txt_id}'})
        nc = ET.SubElement(nc, 'connection',
                           {'class':'textscan','directory':directory,'filename':filename,
                            'password':'','server':''})
        con_1 = ET.SubElement(connection,'_.fcp.ObjectModelEncapsulateLegacy.false...relation',
                              {'connection':f'{txt_id}','name':filename, 'table':f'[{table_name}]',
                               'type':'table'})
        cols_1 = ET.SubElement(con_1, 'columns', {'character-set':'UTF-8','header':'yes','locale':'en_US','separator':','})

        df = pd.read_csv(self.filepath)
        df_types = pd.DataFrame(df.dtypes).reset_index()

        ordinal_1 = create_ordinal(cols_1, df_types)

        con_1 = ET.SubElement(connection,'_.fcp.ObjectModelEncapsulateLegacy.true...relation',
                      {'connection':f'{txt_id}','name':filename, 'table':f'[{table_name}]', 'type':'table'})

        cols_1 = ET.SubElement(con_1, 'columns', {'character-set':'UTF-8','header':'yes','locale':'en_US','separator':','})

        ordinal_2 = create_ordinal(cols_1, df_types)

        metadata_records  = ET.SubElement(connection,'metadata-records')
        mr = ET.SubElement(metadata_records, 'metadata-record', {'class':'capability'})
        remote_name = ET.SubElement(mr, 'remote-name')
        remote_type = ET.SubElement(mr, 'remote-type')
        remote_type.text = str(0)
        parent_name = ET.SubElement(mr, 'parent-name')
        parent_name.text = f"[{filename}]"
        remotealias = ET.SubElement(mr,'remote-alias')
        aggregation = ET.SubElement(mr, 'aggregation')
        aggregation.text = "Count"
        contains_null = ET.SubElement(mr, 'contains-null')
        contains_null.text = "true"
        attributes = ET.SubElement(mr, 'attributes')

        at_1 = ET.SubElement(attributes, 'attribute', {'datatype':'string','name':'character-set'})
        at_1.text = html.escape('"UTF-8"')

        at_2 = ET.SubElement(attributes, 'attribute', {'datatype':'string','name':'collation'})
        at_2.text = html.escape('"en_US"')

        at_3 = ET.SubElement(attributes, 'attribute', {'datatype':'string','name':'field-delimiter'})
        at_3.text = html.escape('","')

        at_4 = ET.SubElement(attributes, 'attribute', {'datatype':'string','name':'header-row'})
        at_4.text = html.escape('"true"')

        at_5 = ET.SubElement(attributes, 'attribute', {'datatype':'string','name':'locale'})
        at_5.text = html.escape('"en_US"')

        at_5 = ET.SubElement(attributes, 'attribute', {'datatype':'string','name':'single-char'})
        at_5.text = html.escape('""')

        ordinal_3 = create_ordinal(metadata_records, df_types, wht="rm", csv_id=csv_id,filename=filename)

        ws_aliases = ET.SubElement(datasource, 'aliases', {'enabled':'yes'})
        
        thisdict = {}
        
        if calculations != None:
            for c in range(0,len(calculations),3):
                cal_caption = calculations[c]
                cal_type = calculations[c+1]
                cal_formula = calculations[c+2]
                cal_name = "Calculation_" + str(c + 1)
                
                thisdict.update({cal_caption : cal_name})
                for key in thisdict.keys():
                    if key in cal_formula:
                        cal_formula = cal_formula.replace(key, thisdict[key])
                
                if cal_type in ("boolean","string"):
                    cal_role = 'dimension'
                    cal_t = 'nominal'
                
                if cal_type in ("real","integer"):
                    cal_role = 'measure'
                    cal_t = 'quantitative'
                                        
                ws_calc = ET.SubElement(datasource, 'column', 
                                        {'caption':cal_caption,
                                         'datatype':cal_type, 
                                         'name':f'[{cal_name}]',
                                         'role':cal_role,
                                         'type':cal_t})
                ws_cals = ET.SubElement(ws_calc, 'calculation', 
                                        {'class':'tableau', 
                                         'formula': cal_formula.replace("&#10;","")
                                        })
        
        if aliases != None:
            for a in range(0, len(aliases),2):
                col_name = aliases[a]
                col_vars = aliases[a+1]
                ws_ali_col = ET.SubElement(datasource, 'column', 
                                           {'datatype':'string', 'name':f'[{col_name}]', 
                                            'role':'dimension', 'type':'nominal'})

                ws_aliases = ET.SubElement(ws_ali_col, 'aliases')
                for col_var in col_vars:
                    ws_alias = ET.SubElement(ws_aliases, 'alias', 
                                             {'key':html.escape(f'"{col_var}"').replace('&quot;','"'), 'value':f'{col_vars[col_var]}'})
        
        col_cap = ET.SubElement(datasource, '_.fcp.ObjectModelTableType.true...column',
                    {'caption':f'{filename}', 'datatype':'table', 
                     'name':f'[__tableau_internal_object_id__].[{csv_id}]',
                     'role':'measure', 'type':'quantitative'})
        
        if drill_path != None:
            for a in range(0, len(drill_path),2):
                drill_paths = ET.SubElement(datasource, 'drill-paths')
                dp = ET.SubElement(drill_paths, 'drill-path', {'name':f'{drill_path[a]}'})
                
                for b in drill_path[a+1]:          
                    dp_field = ET.SubElement(dp, 'field')
                    dp_field.text = f"[{b}]"
            
        col_cap = ET.SubElement(datasource, 'layout',
                                {'_.fcp.SchemaViewerObjectModel.false...dim-percentage':'0.5',
                                 '_.fcp.SchemaViewerObjectModel.false...measure-percentage':'0.4',
                                 'dim-ordering':'alphabetic', 'measure-ordering':'alphabetic', 'show-structure':'true'})
        semantic_values = ET.SubElement(datasource, 'semantic-values')

        semantic_values = ET.SubElement(semantic_values, 'semantic-value', 
                                        {'key':'[Country].[Name]', 'value': html.escape('"India"').replace('&quot;','"')})

        obj_grph = ET.SubElement(datasource, '_.fcp.ObjectModelEncapsulateLegacy.true...object-graph')
        obj = ET.SubElement(obj_grph, 'objects')
        obj = ET.SubElement(obj, 'object', {'caption':f'{filename}', 'id':f'{csv_id}'})
        obj = ET.SubElement(obj, 'properties', {'context':''})
        obj = ET.SubElement(obj, 'relation', 
                            {'connection':f'{txt_id}', 
                             'name':f'{filename}', 'table':f'[{table_name}]', 'type':'table'})
        obj = ET.SubElement(obj, 'columns',  {'character-set':'UTF-8', 'header':'yes', 'locale':'en_US', 'separator':','})

        ordinal_4 = create_ordinal(obj, df_types)
        
        if parameter != None:
            m_datasource = ET.SubElement(datasources, 'datasource',
                                     {'hasconnection':'false', 'inline':'true', 'name':'Parameters', 'version':'18.1'})
            m_aliased = ET.SubElement(m_datasource, 'aliases', {'enabled':'yes'})
                
            for a in range(0, len(parameter),5):
                para_name = parameter[a]
                para_dtype = parameter[a + 1]
                para_type = parameter[a + 2]
                para_value = parameter[a + 3]
                paras = parameter[a + 4]
                
                if para_dtype == 'string': t = 'nominal'
                if para_dtype == 'integer': t = 'quantitative'
                
                paraname = para_name.replace(" ","")
                
                if para_dtype == 'string': 
                    m_col = ET.SubElement(m_datasource, 'column',
                                 {'caption':f'{para_name}', 
                                  'datatype':f'{para_dtype}', 
                                  'name':f'[{paraname}]',
                                  'param-domain-type':f'{para_type}',
                                  'role':'measure', 
                                  'type':f'{t}',
                                  'value': html.escape(f'"{para_value}"').replace('&quot;','"')
                                  })
                    m_cal = ET.SubElement(m_col, 'calculation', 
                                      {'class':'tableau', 
                                       'formula':html.escape(f'"{para_value}"').replace('&quot;','"')})
                    m_members = ET.SubElement(m_col,'members')
                    
                    for b in paras:
                        m_member = ET.SubElement(m_members, 'member', {'value':html.escape(f'"{b}"').replace('&quot;','"')})
                        
                if para_dtype == 'integer':
                    m_col = ET.SubElement(m_datasource, 'column',
                                 {'caption':f'{para_name}', 
                                  'datatype':f'{para_dtype}', 
                                  'name':f'[{paraname}]',
                                  'param-domain-type':f'{para_type}',
                                  'role':'measure',
                                  'type':f'{t}',
                                  'value': f'{para_value}'
                                  })
                    m_cal = ET.SubElement(m_col, 'calculation', 
                                      {'class':'tableau', 
                                       'formula':f'{para_value}'})
                    m_range = ET.SubElement(m_col, 'range', {'granularity':paras[0],
                                                              'max':paras[1],
                                                              'min':paras[2]})
        worksheets = ET.SubElement(workbook, 'worksheets')
        
        for sht in shts:
            ws_name, col_for_cal, what_to_cal, sheet_cnt, a1, a2, apply_filter, \
               customized_label, chart_type, display_field_labels, \
               axisline_visibility, dropline_visibility, refline_visibility, \
               gridline_col_visibility, gridline_row_visibility, zeroline_row_visibility, \
               trendline_visibility, bar_color, text_label, font_family, \
               font_size, font_weight, text_align_h, text_align_v, \
               axis_title, viewpoint, grid_line, color_bars, sheet_cnt, a1, a2, color_var =  sht
            uuid = str(a1) + "E" + str(a2) + "B-66D5-4E6E-A12A-80757EA6A2A8"
            worksheet = ET.SubElement(worksheets, 'worksheet', {'name':f'{ws_name}'})
            ws_table = ET.SubElement(worksheet, 'table')
            ws_view = ET.SubElement(ws_table, 'view')
            ws_datasources = ET.SubElement(ws_view, 'datasources')
            ws_datasource = ET.SubElement(ws_datasources, 'datasource', {'name':'Parameters'})
            
            if color_var != None:
                bar_color = True
            else:
                bar_color = False
            if chart_type == "h":
                scope_c = "cols"
                scope_r = "rows"
                
            if chart_type == "v":
                scope_c = "rows"
                scope_r = "cols"
                
            if chart_type in ('t', 'h','v'):
                ws_datasource = ET.SubElement(ws_datasources, 'datasource',
                                              {'caption':f'{caption_txt}', 'name':f'{fid}'})
                ws_ds_dependencies = ET.SubElement(ws_view, 'datasource-dependencies',
                                                   {'datasource':f'{fid}'})
                for col_for_cals in col_for_cal:
                    ws_col = ET.SubElement(ws_ds_dependencies, 'column',
                                           {'datatype':'string', 'name':f'[{col_for_cals}]', 
                                            'role':'dimension', 'type':'nominal'})
                ws_fcp = ET.SubElement(ws_ds_dependencies, '_.fcp.ObjectModelTableType.false...column',
                                       {'caption':f'{filename}', 'datatype':'integer',
                                        'name':f'[__tableau_internal_object_id__].[{csv_id}]',
                                        'role':'measure', 'type':'quantitative'})
                ws_col = ET.SubElement(ws_ds_dependencies, 'column-instance',
                                       {'column':f'[__tableau_internal_object_id__].[{csv_id}]',
                                        'derivation':'Count',
                                        'name':f'[__tableau_internal_object_id__].[cnt:{csv_id}:qk]',
                                        'pivot':'key',
                                        'type':'quantitative'})

                ws_fcp = ET.SubElement(ws_ds_dependencies, '_.fcp.ObjectModelTableType.true...column',
                                       {'caption':f'{filename}', 'datatype':'table',
                                        'name':f'[__tableau_internal_object_id__].[{csv_id}]',
                                        'role':'measure', 'type':'quantitative'})
                for col_for_cals in col_for_cal:
                    ws_col = ET.SubElement(ws_ds_dependencies, 'column-instance', 
                                           {'column':f'[{col_for_cals}]', 'derivation':'None',
                                            'name':f'[none:{col_for_cals}:nk]', 'pivot':'key', 'type':'nominal'})

            ws_aggregation = ET.SubElement(ws_view, 'aggregation', {'value':'true'})
            ws_style = ET.SubElement(ws_table, 'style')

            if axisline_visibility == False:
                ws_style_rule = ET.SubElement(ws_style, 'style-rule', {'element':'axis'})
                ws_format = ET.SubElement(ws_style_rule, 'format', {'attr':'stroke-size', 'value':'0'})
                ws_format = ET.SubElement(ws_style_rule, 'format', {'attr':'line-visibility', 'value':'off'})
                ws_format = ET.SubElement(ws_style_rule, 'format', {'attr':'tick-color', 'value':'#00000000'})
                ws_format = ET.SubElement(ws_style_rule, 'format', 
                                          {'attr':'display', 'class':'0',
                                           'field': f'[{fid}].[__tableau_internal_object_id__].[cnt:{csv_id}:qk]',
                                           'scope':scope_c, 'value':'false'})
            if display_field_labels == False:
                ws_style_rule = ET.SubElement(ws_style, 'style-rule', {'element':'worksheet'})
                ws_format = ET.SubElement(ws_style_rule, 'format', 
                                          {'attr':'display-field-labels', 'scope':'rows', 'value':'false'})

            if dropline_visibility == False:
                ws_style_rule = ET.SubElement(ws_style, 'style-rule', {'element':'dropline'})
                ws_format = ET.SubElement(ws_style_rule, 'format', {'attr':'stroke-size', 'value':'0'})
                ws_format = ET.SubElement(ws_style_rule, 'format', {'attr':'line-visibility', 'value':'off'})

            if refline_visibility == False:
                ws_style_rule = ET.SubElement(ws_style, 'style-rule', {'element':'refline'})
                ws_format = ET.SubElement(ws_style_rule, 'format', {'attr':'stroke-size', 'value':'0'})
                ws_format = ET.SubElement(ws_style_rule, 'format', {'attr':'line-visibility', 'value':'off'})

            if gridline_col_visibility == False:
                ws_style_rule = ET.SubElement(ws_style, 'style-rule', {'element':'gridline'})
                ws_format = ET.SubElement(ws_style_rule, 'format', {'attr':'stroke-size', 'scope':'cols', 'value':'0'})
                ws_format = ET.SubElement(ws_style_rule, 'format', {'attr':'line-visibility', 'scope':'cols', 'value':'off'})

            if gridline_row_visibility == False:
                ws_style_rule = ET.SubElement(ws_style, 'style-rule', {'element':'gridline'})
                ws_format = ET.SubElement(ws_style_rule, 'format', {'attr':'stroke-size', 'scope':'rows', 'value':'0'})
                ws_format = ET.SubElement(ws_style_rule, 'format', {'attr':'line-visibility','scope':'rows', 'value':'off'})

            if zeroline_row_visibility == False:
                ws_style_rule = ET.SubElement(ws_style, 'style-rule', {'element':'zeroline'})
                ws_format = ET.SubElement(ws_style_rule, 'format', {'attr':'stroke-size', 'value':'0'})
                ws_format = ET.SubElement(ws_style_rule, 'format', {'attr':'line-visibility','value':'off'})

            ws_panes = ET.SubElement(ws_table, 'panes')
            ws_pane = ET.SubElement(ws_panes, 'pane', {'selection-relaxation-option':'selection-relaxation-allow'})
            ws_view = ET.SubElement(ws_pane, 'view')
            ws_breakdown = ET.SubElement(ws_view, 'breakdown', {'value':'auto'})
            ws_mark = ET.SubElement(ws_pane, 'mark', {'class':'Automatic'})

            if trendline_visibility == False:
                ws_style_rule = ET.SubElement(ws_style, 'style-rule', {'element':'trendline'})
                ws_format = ET.SubElement(ws_style_rule, 'format', {'attr':'stroke-size', 'value':'0'})
                ws_format = ET.SubElement(ws_style_rule, 'format', {'attr':'line-visibility','value':'off'})

            if chart_type in ('h','v') and bar_color == True:
                ws_encodings = ET.SubElement(ws_pane, 'encodings')
                ws_color = ET.SubElement(ws_encodings, 'color', {'column':f'[{fid}].[none:{color_var}:nk]'})
                ws_text = ET.SubElement(ws_encodings, 'text', 
                                        {'column':f'[{fid}].[__tableau_internal_object_id__].[cnt:{csv_id}:qk]'})
            if chart_type in ('t'):
                ws_encodings = ET.SubElement(ws_pane, 'encodings')
                ws_text = ET.SubElement(ws_encodings, 'text',
                                        {'column':f'[{fid}].[__tableau_internal_object_id__].[cnt:{csv_id}:qk]'})
                ws_style = ET.SubElement(ws_pane, 'style')
                ws_style_rule = ET.SubElement(ws_style, 'style-rule', {'element':'mark'})
                ws_format = ET.SubElement(ws_style_rule, 'format', 
                                          {'attr':'mark-labels-show', 'value':'true'})

            if text_label == True and chart_type in ('h','v'):
                wd_style = ET.SubElement(ws_pane, 'style')
                wd_style_rule = ET.SubElement(wd_style, 'style-rule', {'element':'mark'})
                wd_format = ET.SubElement(wd_style_rule, 'format', {'attr':'mark-labels-show', 'value':'true'})
                wd_format = ET.SubElement(wd_style_rule, 'format', {'attr':'mark-labels-cull', 'value':'true'})

            ws_rows = ET.SubElement(ws_table,'rows')
            if chart_type in ('t', 'h'):
                x =0
                for col_for_cals in col_for_cal:
                    if x==0:
                        row_txt = f'[{fid}].[none:{col_for_cals}:nk]'
                    else:
                        row_txt = row_txt + " / " + f'[{fid}].[none:{col_for_cals}:nk]'
                x +=1
                
                ws_rows.text = row_txt
                    
            if chart_type in ('v'):
                ws_rows.text = f'[{fid}].[__tableau_internal_object_id__].[cnt:{csv_id}:qk]'

            ws_cols = ET.SubElement(ws_table,'cols')
            if chart_type in ("h"):
                ws_cols.text = f'[{fid}].[__tableau_internal_object_id__].[cnt:{csv_id}:qk]'
            if chart_type in ("v"):
                x = 0
                for col_for_cals in col_for_cal:
                    if x==0:
                        row_txt = f'[{fid}].[none:{col_for_cals}:nk]'
                    else:
                        row_txt = row_txt + " / " + f'[{fid}].[none:{col_for_cals}:nk]'
                    x +=1
                ws_cols.text = row_txt

            ws_simple_id = ET.SubElement(worksheet, 'simple-id', {'uuid':'{' + f'{uuid}' + '}'})

# Windows
        window = ET.SubElement(workbook,'windows', {'source-height':'30'})
        for sht in shts:
            ws_name, col_for_cal, what_to_cal, sheet_cnt, a1, a2, apply_filter, \
               customized_label, chart_type, display_field_labels, \
               axisline_visibility, dropline_visibility, refline_visibility, \
               gridline_col_visibility, gridline_row_visibility, zeroline_row_visibility, \
               trendline_visibility, bar_color, text_label, font_family, \
               font_size, font_weight, text_align_h, text_align_v, \
               axis_title, viewpoint, grid_line, color_bars, sheet_cnt, a1, a2, color_var =  sht
            
            uuid = str(a1) + "E" + str(a2) + "B-66D5-4E6E-A12A-80757EA6A2A8"
            wd = ET.SubElement(window, 'window', {'class':'worksheet', 'name':f'{ws_name}'})
            wd_cards = ET.SubElement(wd, 'cards')
            wd_edge = ET.SubElement(wd_cards, 'edge', {'name':'left'})
            wd_strip = ET.SubElement(wd_edge, 'strip', {'size':'160'})
            wd_card = ET.SubElement(wd_strip, 'card', {'type':'pages'})
            wd_card = ET.SubElement(wd_strip, 'card', {'type':'filters'})
            wd_card = ET.SubElement(wd_strip, 'card', {'type':'marks'})
            wd_edge = ET.SubElement(wd_cards, 'edge', {'name':'top'})
            wd_strip = ET.SubElement(wd_edge, 'strip', {'size':'2147483647'})
            wd_card = ET.SubElement(wd_strip, 'card', {'type':'columns'})
            wd_strip = ET.SubElement(wd_edge, 'strip', {'size':'2147483647'})
            wd_card = ET.SubElement(wd_strip, 'card', {'type':'rows'})
            wd_strip = ET.SubElement(wd_edge, 'strip', {'size':'31'})
            wd_card = ET.SubElement(wd_strip, 'card', {'type':'title'})
            if bar_color == True:
                wd_edge = ET.SubElement(wd_cards, 'edge', {'name':'right'})
                wd_strip = ET.SubElement(wd_edge, 'strip', {'size':'160'})
                wd_card = ET.SubElement(wd_strip, 'card', 
                                        {
                                            'pane-specification-id':'0',
                                             'param':f'[{fid}].[none:{col_for_cal}:nk]',
                                             'type':'color'
                                        }
                                       )

            wd_viewpoint = ET.SubElement(wd, 'viewpoint')

            if viewpoint != 'standard':
                wd_zoom = ET.SubElement(wd_viewpoint, 'zoom', {'type':f'{viewpoint}'})

            if bar_color == True:
                wd_highlight = ET.SubElement(wd_viewpoint, 'highlight')
                wd_color_one_way = ET.SubElement(wd_highlight, 'color-one-way')
                
                for col_for_cals in col_for_cal:
                    wd_field = ET.SubElement(wd_color_one_way, 'field')
                    wd_field.text = f'[{fid}].[none:{col_for_cals}:nk]'
                    
            wd_simple_id = ET.SubElement(wd, 'simple-id', {'uuid':'{' + f'{uuid}' + '}'})
    
        tree = ET.ElementTree(workbook)
        tree.write(output_name, encoding='utf-8', xml_declaration=True)
        
        print("Tableau workbook created successfully.")
        
tab = tableau(filepath="C:/Users/g646787/Downloads/03_Python Session/python_code/ocd.csv",
              aliases = ["Super_Category", 
                         {'alternative_breakfast':'ALTERNATIVE BREAKFAST', 'baking':'BAKING', 'pet':"PET"}],
              drill_path = ["Product Hierarchy", 
                            ['Super_Category','Category','Manufacturer','Higher_Level_Themes','Clusters','Sub_Clusters']],
              parameter = [
                  'Para_Select_Flag', 'string', 'list', 'Matured', ['Emerging','Matured','Small growing trends'],
                  'Para_X_Axis', 'string', 'list', '2 Yr CAGR Sales', ['2 Yr CAGR Sales',
                                                                'L52W Growth Sales',
                                                                'L26W Growth Sales',
                                                                'L13W Growth Sales',
                                                                '2Yr CAGR EQ',
                                                                'L52W EQ Growth',
                                                                'L26W EQ Growth',
                                                                'L13W EQ Growth'
                                                               ],
                  'Para_Y_Axis','string','list','2 Yr TDP CAGR',['2 Yr TDP CAGR',
                                                                 'L52W TDP Growth',
                                                                 'L26W TDP Growth',
                                                                 'L13W TDP Growth'
                                                                ],
                  'Para_Top','integer','range', '1', ['1','15','1']
                           ],
              calculations = ["@X_Axis_Calc", 'real',
                              """IF [Parameters].[Para_X_Axis]="3 Yr CAGR Sales" THEN [Sales_3yr_CAGR] &#10;
                                     ELSEIF [Parameters].[Para_X_Axis]="2 Yr CAGR Sales" THEN [Sales_2yr_CAGR] &#10;
                                     ELSEIF [Parameters].[Para_X_Axis]="L52W Growth Sales" THEN [52_Weeks_Sales_Growth] &#10;
                                     ELSEIF [Parameters].[Para_X_Axis]="L26W Growth Sales" THEN [26_Weeks_Sales_Growth] &#10;
                                     ELSEIF [Parameters].[Para_X_Axis]="L13W Growth Sales" THEN [13_Weeks_Sales_Growth] &#10;
                                     ELSEIF [Parameters].[Para_X_Axis]="L52W EQ Growth" THEN [52_Weeks_eq_vol_Growth] &#10;
                                     ELSEIF [Parameters].[Para_X_Axis]="L26W EQ Growth" THEN [26_Weeks_eq_vol_Growth] &#10;
                                     ELSEIF [Parameters].[Para_X_Axis]="L13W EQ Growth" THEN [13_Weeks_eq_vol_Growth] &#10;
                                     ELSEIF [Parameters].[Para_X_Axis]="3 Yr EQ CAGR" THEN [eq_vol_3yr_CAGR] &#10;
                                     ELSEIF [Parameters].[Para_X_Axis]="2 Yr EQ CAGR" THEN [eq_vol_2yr_CAGR] &#10;
                                END""",
                              "@Y_Axis_Calc", 'real',
                              """IF [Parameters].[Para_Y_Axis] = "L52W TDP Growth" THEN [52_Weeks_tdp_Growth] &#10;
                                    ELSEIF [Parameters].[Para_Y_Axis] = "L26W TDP Growth" THEN [26_Weeks_tdp_Growth] &#10;
                                    ELSEIF [Parameters].[Para_Y_Axis] = "L13W TDP Growth" THEN [13_Weeks_tdp_Growth] &#10;
                                    ELSEIF [Parameters].[Para_Y_Axis] = "2 Yr TDP CAGR" THEN [tdp_2yr_CAGR] &#10;
                                    ELSEIF [Parameters].[Para_Y_Axis] = "3 Yr TDP CAGR" THEN [tdp_3yr_CAGR] &#10;
                                END""",
                              "@X_Axis_Ref_Line", 'real',
                              """IF [Parameters].[Para_X_Axis] ="3 Yr CAGR Sales" THEN [Category_Sales_3yr_CAGR] &#10;
                                    ELSEIF [Parameters].[Para_X_Axis] ="2 Yr CAGR Sales" THEN [Category_Sales_2yr_CAGR] &#10;
                                    ELSEIF [Parameters].[Para_X_Axis] ="L52W Growth Sales"  then [Category_52_Weeks_Sales_Growth] &#10;
                                    ELSEIF [Parameters].[Para_X_Axis] ="L26W Growth Sales" then [Category_26_Weeks_Sales_Growth] &#10;
                                    ELSEIF [Parameters].[Para_X_Axis] ="L13W Growth Sales" then [Category_13_Weeks_Sales_Growth] &#10;
                                    ELSEIF [Parameters].[Para_X_Axis] ="L52W EQ Growth" THEN [Category_52_Weeks_eq_vol_Growth] &#10;
                                    ELSEIF [Parameters].[Para_X_Axis] ="L26W EQ Growth" THEN [Category_26_Weeks_eq_vol_Growth] &#10;
                                    ELSEIF [Parameters].[Para_X_Axis] ="L13W EQ Growth" THEN [Category_13_Weeks_eq_vol_Growth] &#10;
                                    ELSEIF [Parameters].[Para_X_Axis] ="3 Yr EQ CAGR" THEN [Category_eq_vol_3yr_CAGR] &#10;
                                    ELSEIF [Parameters].[Para_X_Axis] ="2 Yr EQ CAGR" THEN [Category_eq_vol_2yr_CAGR] &#10;
                                END""",
                              "@Rank",'integer',"RANK(SUM([@X_Axis_Calc]),'desc')",
                              "@Select_Clusters",'boolean',"[@Rank] <= [Parameters].[Para_Top]",
                              "@Cluster_Title",'string',"'Cluster Performance'",
                              "@Select_Flag",'boolean',"[Flag]=[Parameters].[Para_Select_Flag]"
                             ]
             )

sht_1 = tab.sheets(ws_name = "Sheet_Test_1", 
                   col_for_cal = ["@X_Axis_Calc"], 
                   what_to_cal="Count",
                   display_field_labels = False
                  )
tab.create_dashboard('C:/Users/g646787/Downloads/03_Python Session/python_code/tableau_testing.twb', 
                     [sht_1])
