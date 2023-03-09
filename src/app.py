#IMPORT LIBRARY
from dash import Dash, dcc, html, Input, Output, dash_table
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

#CONFIG
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.config['suppress_callback_exceptions']= True
server = app.server

#SET VARIABLE
df = pd.read_excel('./ESPBR.xlsx', sheet_name = 'All')
colors = {'primary': '#66ff66', 'secondary': '#C5FFBA', 'background': '#F3F6FA'}
CONST_TGT_ENQ = 200
CONST_TGT_SUS = 25
CONST_TGT_PROS = 10
CONST_TGT_SPK = 3
CONST_TGT_DO = 2


#FUNCTION
def checking(x):
    arr = []
    for elements in x:
        if elements >= 0:
            arr.append('YES')
        else:
            arr.append('')
    return arr

def totalChecking(x):
    cek = 0
    for elements in x:
        if elements == 'YES':
            cek = cek + 1
    
    return cek

def percentage(x, y):
    arr = []
    for elements in x:
        arr.append('{}%'.format(round((elements / y * 100))))

    return arr

def createGraph(y_axes, x_axes, tgt, legend, target_legend):
    fig = go.Figure()

    fig.add_trace(go.Bar(
        y=y_axes,
        x=tgt,
        orientation='h',
        name=target_legend,
        hoverinfo = 'none',
        marker=dict(color = colors['secondary'])
    ))
    fig.add_trace(go.Bar(
        y=y_axes,
        x=x_axes,
        orientation='h',
        name=legend,
        text=x_axes,
        hoverinfo = 'none',
        marker=dict(color = colors['primary'])
    ))         
    fig.update_layout(barmode='overlay',
                      plot_bgcolor= 'rgba(0, 0, 0, 0)',
                      yaxis = dict(
                        tickfont = dict(size = 10)
                      ),
                      legend=dict(
                      yanchor="bottom",
                      y=1.02,
                      xanchor="right",
                      x=1,
                      orientation='h'
                      ),
                      margin=dict(l=10, r=10, t=20, b=20),
                      autosize=True,
                      height=400,
                      )
    return fig

def emptyGraph():
    return {
        "layout": {
            "xaxis": {
                "visible": False
            },
            "yaxis": {
                "visible": False
            },
            "annotations": [
                {
                    "text": "No matching data found",
                    "xref": "paper",
                    "yref": "paper",
                    "showarrow": False,
                    "font": {
                        "size": 28
                    }
                }
            ]
        }
    }

def createTable (data_table): 
    table = dash_table.DataTable(data=data_table.to_dict('records'),
                                columns=[
                                    {'name': ['', 'Sales'], 'id': 'Sales'},
                                    {'name': ['Enquiry', 'Tgt'], 'id': 'Tgt Enq'},
                                    {'name': ['Enquiry', 'Acv'], 'id': 'E'},
                                    {'name': ['Enquiry', 'Diff'], 'id': 'Diff_E'},
                                    {'name': ['Enquiry', '%'], 'id': 'persen_E'},
                                    {'name': ['Enquiry', 'Check'], 'id': 'enq_check'},
                                    {'name': ['Suspect', 'Tgt'], 'id': 'Tgt Sus'},
                                    {'name': ['Suspect', 'Acv'], 'id': 'S'},
                                    {'name': ['Suspect', 'Diff'], 'id': 'Diff_S'},
                                    {'name': ['Suspect', '%'], 'id': 'persen_S'},
                                    {'name': ['Suspect', 'Check'], 'id': 'sus_check'}, 
                                    {'name': ['Prospect', 'Tgt'], 'id': 'Tgt Pros'},
                                    {'name': ['Prospect', 'Acv'], 'id': 'P'},
                                    {'name': ['Prospect', 'Diff'], 'id': 'Diff_P'},
                                    {'name': ['Prospect', '%'], 'id': 'persen_P'},
                                    {'name': ['Prospect', 'Check'], 'id': 'pros_check'},
                                    {'name': ['SPK', 'Tgt'], 'id': 'Tgt SPK'},
                                    {'name': ['SPK', 'Acv'], 'id': 'SPK'},
                                    {'name': ['SPK', 'Diff'], 'id': 'Diff_SPK'},
                                    {'name': ['SPK', '%'], 'id': 'persen_spk'},
                                    {'name': ['SPK', 'Check'], 'id': 'spk_check'},
                                    {'name': ['DO', 'Tgt'], 'id': 'Tgt DO'},
                                    {'name': ['DO', 'Acv'], 'id': 'DO'},
                                    {'name': ['DO', 'Diff'], 'id': 'Diff_DO'},
                                    {'name': ['DO', '%'], 'id': 'persen_do'},
                                    {'name': ['DO', 'Check'], 'id': 'do_check'},                                       
                                ],
                                style_table={'overflowX': 'auto', 'minWidth': '100%', 'width':'50%'},
                                sort_action="native",
                                sort_mode="multi",
                                merge_duplicate_headers=True,
                                style_data={
                                    'textAlign': 'center'
                                },
                                style_header={
                                    'textAlign': 'center'
                                },
                                style_header_conditional=[{
                                    'if':{'header_index': 0},
                                    'backgroundColor': colors['primary']
                                },
                                {
                                    'if':{'header_index': 1},
                                    'backgroundColor': colors['secondary'],
                                }],

                                style_data_conditional=[{
                                    'if':{'column_id': c, 'filter_query': '{' + c + '}' + ' < 0'},
                                    'color': 'red'
                                } for c in data_table.columns],

                                fixed_columns={'headers': True,'data': 1},
                         
                                )
    return table

def createDataFrame(group, month, year):
    df1 = df.loc[(df['Grup'] == group) & (df['Bulan'] == month) & (df['Tahun'] == year)]
    total = df1.groupby(['Sales'], as_index = False).sum()
    
    total['Tgt Enq'] = CONST_TGT_ENQ
    total['Tgt Sus'] = CONST_TGT_SUS
    total['Tgt Pros'] = CONST_TGT_PROS
    total['Tgt SPK'] = CONST_TGT_SPK
    total['Tgt DO'] = CONST_TGT_DO

    diff_enq = total['E'] - total['Tgt Enq']
    diff_sus = total['S'] - total['Tgt Sus']
    diff_pros = total['P'] - total['Tgt Pros']
    diff_SPK = total['SPK'] - total['Tgt SPK']
    diff_DO = total['DO'] - total['Tgt DO']

    persen_enq = percentage(total['E'].tolist(), CONST_TGT_ENQ)
    persen_sus = percentage(total['S'].tolist(), CONST_TGT_SUS)
    persen_pros = percentage(total['P'].tolist(), CONST_TGT_PROS)
    persen_SPK = percentage(total['SPK'].tolist(), CONST_TGT_SPK)
    persen_DO = percentage(total['DO'].tolist(), CONST_TGT_DO)

    check_enq = checking(diff_enq)
    check_sus = checking(diff_sus)
    check_pros = checking(diff_pros)
    check_SPK = checking(diff_SPK)
    check_DO = checking(diff_DO)

    total = total.assign(Diff_E= diff_enq, Diff_S = diff_sus, Diff_P = diff_pros,
                         Diff_SPK = diff_SPK, Diff_DO = diff_DO, persen_E = persen_enq,
                         persen_S = persen_sus, persen_P = persen_pros, persen_spk = persen_SPK,
                         persen_do = persen_DO, enq_check = check_enq, sus_check = check_sus,
                         pros_check = check_pros, spk_check = check_SPK, do_check = check_DO)

    total= total[['Sales', 'Tgt Enq', 'E', 'Diff_E', 'persen_E', 'enq_check', 
                  'Tgt Sus', 'S', 'Diff_S', 'persen_S', 'sus_check',
                  'Tgt Pros', 'P', 'Diff_P', 'persen_P', 'pros_check',
                  'Tgt SPK', 'SPK', 'Diff_SPK', 'persen_spk', 'spk_check',
                  'Tgt DO', 'DO', 'Diff_DO', 'persen_do', 'do_check',]]

    return total

def createCard(aaa, bbb):
    card = dbc.Card([
        dbc.CardBody(
            [
                html.H5(id = aaa),
                html.P(bbb),
            ]
        )
    ], color=colors['primary'], className="border-0")

    return card

def esp(datas, var,const_tgt, type, target_type):
    y_axes = datas['Sales'].tolist()
    x_axes = datas[var].tolist()
    tgt = []        

    for i in range(len(y_axes)):
        tgt.append(const_tgt)
    fig = createGraph(y_axes, x_axes, tgt, type, target_type)
    tgt = []

    return fig

def returnItem(data, var, tgt,str_legend, str_name, trgt_name, diff, check):
    fig = esp(data, var, tgt, str_legend, trgt_name)
    percentage_int = round(data[var].sum() / (tgt * len(data[str_name])) * 100)
    tot_acv_card = str(data[var].sum())
    tot_diff_card = str(data[diff].sum())
    tot_percentage_card = '{}%'.format(percentage_int)
    tot_total_check = str(totalChecking(data[check]))

    return fig, tot_acv_card, tot_diff_card, tot_percentage_card, tot_total_check  

#LAYOUT
app.layout = html.Div([
    dbc.Row([   
        html.H2('Performance Sales Honda Sonic'),
        html.H4(id = 'month_year'),
        html.H5(id = 'group_header'),
        html.Hr()
    ], className = 'text-center'),

    dbc.Row([
        dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.Div([
                        html.P('Month :', className = 'mt-4 mb-2'),
                        dcc.Dropdown(id = 'month_dropdown',
                            options = df['Bulan'].unique(),
                            value = 'Jan'),
                        html.P('Year :', className = 'mt-4 mb-2'),
                        dcc.Dropdown(id = 'year_dropdown',
                            options = df['Tahun'].unique(),
                            value = 2022),
                        html.P('Choose Group : ', className = 'mt-4 mb-2'),
                        dbc.Col(
                            dcc.Dropdown(id = 'group_dropdown',
                            options = df['Grup'].unique(),
                            value = 'BANGUN SUBANGSA'),
                        ),
                        html.P('Select : ', className = 'mt-4 mb-2'),
                        html.Div([
                            dcc.RadioItems(
                            id = 'radio',
                            options=[
                                {'label':' Enquiry', 'value': 'enq'},
                                {'label':' Suspect', 'value': 'sus'},
                                {'label':' Prospect', 'value': 'pros'},
                                {'label':' SPK', 'value': 'spk'},
                                {'label':' DO', 'value': 'do'}
                            ], value='enq', labelStyle={'width':'100%', 'marginBottom': '5px'}
                            )
                        ])
                    ])
                                        
                ])
        ], className='border-0 h-100 bg-white')
        ], width = 3),
        dbc.Col([
            dbc.Row([
                dbc.Col(createCard('acv_card', 'Total Activity')),
                dbc.Col(createCard('diff_card', 'Total Diffrence')),               
                dbc.Col(createCard('percentage_card', 'Total Percentage')),
                dbc.Col(createCard('check_card', 'Total Check'))               
            ], className = 'mb-3'),
            dbc.Card([
                dbc.CardBody([
                    html.H3(id = 'graph_title', className = 'text-center'),
                    dcc.Graph(id = 'my_graph',
                                config= dict(
                                displayModeBar = False
                                )
                    )
                ])
            ], className='border-0 bg-white')
        ], width = 9)
    ]),

    dbc.Row([
        dbc.Card([
            dbc.CardBody([
                html.Div(id = 'my_table')
            ])
        ], className='border-0 bg-white')
    ], className = 'mt-3')
], className = 'px-4 py-2', style={'backgroundColor': colors['background']})


#CALLBACK
@app.callback(Output('my_table', 'children'),
              Output('my_graph', 'figure'),
              Output('acv_card','children'),
              Output('diff_card','children'),
              Output('percentage_card','children'),
              Output('check_card','children'),
              Output('graph_title','children'),
              Input('radio', 'value'),
              Input('group_dropdown', 'value'),
              Input('month_dropdown', 'value'),
              Input('year_dropdown', 'value'))
def update_enq(rad, group, month, year):
    df1 = createDataFrame(group, month, year)
    if (len(df1) == 0):
        fig = emptyGraph()
        table = ''
        tot_acv_card = '-'
        tot_diff_card = '-'
        tot_percentage_card = '-'
        tot_total_check = '-'
        title = ''
    else :
        table = createTable(df1)
        if rad == 'enq':
            fig, tot_acv_card, tot_diff_card, tot_percentage_card, tot_total_check = returnItem(df1, 'E', CONST_TGT_ENQ, 'Enquiry', 'Sales', 'Target Enquiry', 'Diff_E', 'enq_check')
            title = 'Enquiry'

        elif rad == 'sus':
            fig, tot_acv_card, tot_diff_card, tot_percentage_card, tot_total_check = returnItem(df1, 'S', CONST_TGT_SUS, 'Suspect', 'Sales', 'Target Suspect', 'Diff_S', 'sus_check')
            title = 'Suspect'

        elif rad == 'pros':
            fig, tot_acv_card, tot_diff_card, tot_percentage_card, tot_total_check = returnItem(df1, 'P', CONST_TGT_PROS, 'Prospect', 'Sales', 'Target Prospect', 'Diff_P', 'pros_check')
            title = 'Prospect'

        elif rad == 'spk':
            fig, tot_acv_card, tot_diff_card, tot_percentage_card, tot_total_check = returnItem(df1, 'SPK', CONST_TGT_SPK, 'SPK', 'Sales', 'Target SPK', 'Diff_SPK', 'spk_check')
            title = 'SPK'

        elif rad == 'do':
            fig, tot_acv_card, tot_diff_card, tot_percentage_card, tot_total_check = returnItem(df1, 'DO', CONST_TGT_DO, 'DO', 'Sales', 'Target DO', 'Diff_DO', 'do_check')
            title = 'DO'

    return table, fig, tot_acv_card, tot_diff_card, tot_percentage_card, tot_total_check, title

@app.callback(Output('month_year', 'children'),
              Output('group_header', 'children'),
              Input('month_dropdown', 'value'),
              Input('year_dropdown', 'value'),
              Input('group_dropdown', 'value'))
def update_header(month, year, group):
    date = '{} {}'.format(month, year)
    group_header = str(group)

    return date, group_header



#RUN SERVER
if __name__ == "__main__":
    app.run_server(debug=False)