from dash import Dash, html, dash_table, dcc, callback, Output, Input
import pandas as pd
import plotly.express as px
from pathlib import Path


DATE = '2023-08-23'
DATA_DIR = Path(__file__).parent.parent / "data"
EMAILS_FILE_NAME = "ML-" + DATE + ".xlsx"
DEMANDES_FILE_NAME = 'DM-' + DATE + ".xlsx"
EMAILS_FILE = DATA_DIR / EMAILS_FILE_NAME
DEMANDES_FILE = DATA_DIR / DEMANDES_FILE_NAME

def report():
    emails = pd.read_excel(EMAILS_FILE).dropna(how='all')
    emails['Date'] = pd.to_datetime(emails['Date'], utc=True, format="mixed").dt.tz_localize(None)
    emails['Jour'] = emails['Date'].dt.strftime('%a, %d %b %Y')
    emails['Jour'] = pd.to_datetime(emails['Jour'])
    emails_cut = emails[["Thread", "Date", "Jour", "From", "To", "Subject"]]

    # Incorporate data
    demandes = pd.read_excel(DEMANDES_FILE, sheet_name='All').dropna(how='all')
    informations = pd.read_excel(DEMANDES_FILE, sheet_name='Information').dropna(how='all')
    annulations = pd.read_excel(DEMANDES_FILE, sheet_name='Annulation').dropna(how='all')
    erreurs = pd.read_excel(DEMANDES_FILE, sheet_name='Erreur').dropna(how='all')
    activations = pd.read_excel(DEMANDES_FILE, sheet_name='Activation').dropna(how='all')
    creations = pd.read_excel(DEMANDES_FILE, sheet_name='Creation').dropna(how='all')
    modifications = pd.read_excel(DEMANDES_FILE, sheet_name='Modification').dropna(how='all')
    extractions = pd.read_excel(DEMANDES_FILE, sheet_name='Extraction').dropna(how='all')
    alls = pd.concat([informations, annulations, erreurs, activations, creations, modifications, extractions],join='inner', ignore_index=True)

    applications = pd.concat([informations[['Application', 'Type']], annulations[['Application', 'Type']], erreurs[['Application', 'Type']], activations[['Application', 'Type']], creations[['Application', 'Type']], modifications[['Application', 'Type']]])

    # Initialize the app
    app = Dash(__name__)
    colors = {
            'Extraction': 'green',
            'Erreur': 'red',
            'Activation': 'grey',
            'Annulation': 'cyan',
            'Information': 'blue',
            'Creation': 'purple',
            'Modification': 'orange',
            }
    app_colors = {
            'UGOUV': 'grey',
            'UNIV': 'orange',
            'HOSIX': 'blue',
            'LPD': 'green',
            }
    demandes_date_box = px.box(demandes[demandes.Temps != 0], x="Temps", y="Type", color="Type", color_discrete_map=colors)
    demandes_date_box.update_traces(quartilemethod="exclusive") # or "inclusive", or "linear" by default
    demandes_date_box.update_xaxes(tick0=0, dtick=24)
    demandes_date_box.update_xaxes(scaleratio = 1, gridcolor='black', griddash='dash', minor_griddash="dot")
    app.layout = html.Div([
        dash_table.DataTable(
            columns=[{"name": i, "id": i} for i in emails_cut.columns],
            data=emails_cut.to_dict('records'),
            page_size=10,
            style_data={
                'whiteSpace': 'normal',
                'height': 'auto',
                'backgroundColor': 'white',
                'color': 'black'
            },
            style_header={
                'backgroundColor': 'rgb(30, 30, 30)',
                'color': 'white'
            },
        ),
        dcc.Graph(figure=px.histogram(emails_cut, x="Jour", text_auto=True, nbins=(emails_cut.Jour.max() - emails_cut.Jour.min()).days + 1).update_xaxes(dtick=86400000.0).update_layout(xaxis_tickformat = '%a, %d %Y')),
        html.H2('Inforamtions générales'),
        dash_table.DataTable(
            columns=[{"name": i, "id": i} for i in ["Date", "Expediteurs", "Destinataires", "Type", "Objet", "Application", "Temps"]],
            data=alls[["Date", "Expediteurs", "Destinataires", "Type", "Objet", "Application", "Temps"]].to_dict('records'),
            page_size=10,
            style_data={
                'whiteSpace': 'normal',
                'height': 'auto',
                'backgroundColor': 'white',
                'color': 'black'
            },
            style_header={
                'backgroundColor': 'rgb(30, 30, 30)',
                'color': 'white'
            },
        ),
        dcc.Graph(figure=px.scatter(alls, x="Date", y="Objet", color="Application").update_traces(marker_size=10).update_xaxes(dtick=86400000.0, gridcolor='black', minor_griddash="dot")),

        dcc.Graph(figure=px.histogram(demandes, x='Type', color="Type", color_discrete_map=colors, text_auto=True).update_xaxes(categoryorder='total ascending')),
        dcc.Graph(figure=demandes_date_box),
        dcc.Graph(figure=px.histogram(applications, x='Application', color="Type", color_discrete_map=colors, text_auto=True).update_xaxes(categoryorder='total ascending')),
        # 
        html.H2("Extractions"),
        dash_table.DataTable(
            columns=[{"name": i, "id": i} for i in ["Date", "Expediteurs", "Destinataires", "Type", "Objet", "Application", "Temps"]],
            data=alls[["Date", "Expediteurs", "Destinataires", "Type", "Objet", "Application", "Temps"]].to_dict('records'),
            page_size=10,
            style_data={
                'whiteSpace': 'normal',
                'height': 'auto',
                'backgroundColor': 'white',
                'color': 'black'
            },
            style_header={
                'backgroundColor': 'rgb(30, 30, 30)',
                'color': 'white'
            },
        ),
        dcc.Graph(figure=px.pie(extractions, names='Objet', labels='Objet').update_traces(textinfo='percent+label')),
        # dcc.Graph(figure=px.histogram(extractions, x='Demandeur'), id="graph"),
        dcc.Graph(figure=px.scatter(extractions[extractions.Temps != 0], x='Temps').update_xaxes(tick0=0, dtick=24, scaleratio = 1, gridcolor='black', griddash='dash', minor_griddash="dot")),
        dcc.Graph(figure=px.line(extractions, y='Heure', color='Objet', markers=True).update_xaxes()),

        html.H2("Modifications"),
        dash_table.DataTable(
            columns=[{"name": i, "id": i} for i in ["Date", "Type", "Objet", "Action", "Status", "Application", "Demandeur"]],
            data=modifications[["Date", "Type", "Objet", "Action", "Status", "Application", "Demandeur"]].to_dict('records'),
            page_size=10,
            style_data={
                'whiteSpace': 'normal',
                'height': 'auto',
                'backgroundColor': 'white',
                'color': 'black'
            },
            style_header={
                'backgroundColor': 'rgb(30, 30, 30)',
                'color': 'white'
            },
        ),
        dcc.Graph(figure=px.pie(modifications, names='Application', color="Application", color_discrete_map=app_colors).update_traces(textinfo='value+label').update_xaxes(categoryorder='total ascending')),
        dcc.Graph(figure=px.pie(modifications, names='Status', color="Status").update_traces(textinfo='value+label').update_xaxes(categoryorder='total ascending')),
        dash_table.DataTable(
            columns=[{"name": i, "id": i} for i in ["Date", "Type", "Objet", "Action", "Status", "Demandeur"]],
            data=modifications[(modifications["Application"] == "UGOUV") & (modifications["Status"] == "Non fait")][["Date", "Type", "Objet", "Action", "Status", "Demandeur"]].to_dict('records'),
            # page_size=10,
            style_data={
                'whiteSpace': 'normal',
                'height': 'auto',
                'backgroundColor': 'white',
                'color': 'black'
            },
            style_header={
                'backgroundColor': 'rgb(30, 30, 30)',
                'color': 'white'
            },
        ),
        html.H3("UGOUV"),
        dcc.Graph(figure=px.histogram(modifications[(modifications["Application"] == "UGOUV") & (modifications["Status"] == "Fait")], x='Objet', text_auto=True).update_xaxes(categoryorder='total ascending')),
        dash_table.DataTable(
            columns=[{"name": i, "id": i} for i in ["Date", "Type", "Objet", "Action", "Status", "Demandeur"]],
            data=modifications[(modifications["Application"] == "UGOUV") & (modifications["Status"] == "Fait")][["Date", "Type", "Objet", "Action", "Status", "Demandeur"]].to_dict('records'),
            page_size=10,
            style_data={
                'whiteSpace': 'normal',
                'height': 'auto',
                'backgroundColor': 'white',
                'color': 'black'
            },
            style_header={
                'backgroundColor': 'rgb(30, 30, 30)',
                'color': 'white'
            },
        ),
        dcc.Graph(figure=px.scatter(modifications[(modifications.Temps != 0) & (modifications["Status"] == "Fait")], y="Temps", x="Objet", color="Objet").update_traces(marker_size=10).update_yaxes(tick0=0, dtick=24, gridcolor='black', minor_griddash="dot").update_xaxes(gridcolor='black', minor_griddash="dot")),
        dcc.Graph(figure=px.histogram(modifications[(modifications["Application"] == "UGOUV") & (modifications["Status"] == "Fait")], x='Demandeur', color='Objet', text_auto=True).update_xaxes(categoryorder='total ascending')),
        html.H3("Autres applications"),
        dash_table.DataTable(
            columns=[{"name": i, "id": i} for i in ["Date", "Type", "Objet", "Action", "Status", "Demandeur"]],
            data=modifications[(modifications["Application"] != "UGOUV") & (modifications["Status"] == "Fait")][["Date", "Type", "Objet", "Action", "Status", "Demandeur"]].to_dict('records'),
            # page_size=10,
            style_data={
                'whiteSpace': 'normal',
                'height': 'auto',
                'backgroundColor': 'white',
                'color': 'black'
            },
            style_header={
                'backgroundColor': 'rgb(30, 30, 30)',
                'color': 'white'
            },
        ),
        dcc.Graph(figure=px.histogram(modifications[(modifications["Application"] != "UGOUV") & (modifications["Status"] == "Fait")], x='Objet', color='Application', color_discrete_map=app_colors, text_auto=True).update_xaxes(categoryorder='total ascending')),
        dcc.Graph(figure=px.histogram(modifications[(modifications["Application"] != "UGOUV") & (modifications["Status"] == "Fait")], x='Demandeur', color='Objet', text_auto=True).update_xaxes(categoryorder='total ascending')),

        html.H2("Creation"),
        dash_table.DataTable(
            columns=[{"name": i, "id": i} for i in ["Date", "Type", "Objet", "Application", "Demandeur", "Temps"]],
            data=creations[["Date", "Type", "Objet", "Application", "Demandeur", "Temps"]].to_dict('records'),
            page_size=10,
            style_data={
                'whiteSpace': 'normal',
                'height': 'auto',
                'backgroundColor': 'white',
                'color': 'black'
            },
            style_header={
                'backgroundColor': 'rgb(30, 30, 30)',
                'color': 'white'
            },
        ),

        dcc.Graph(figure=px.pie(creations, names='Application', color="Application", color_discrete_map=app_colors).update_traces(textinfo='value+label')),
        dcc.Graph(figure=px.histogram(creations, x='Application', color='Objet', text_auto=True).update_xaxes(categoryorder='total ascending')),
        dcc.Graph(figure=px.histogram(creations, x='Objet', color='Demandeur').update_xaxes(categoryorder='total ascending')),
        dcc.Graph(figure=px.scatter(creations, x="Date", y="Objet", color="Application").update_traces(marker_size=10).update_xaxes(dtick=86400000.0, gridcolor='black', minor_griddash="dot")),
        # dcc.Graph(figure=px.box(creations[creations.Temps != 0]['Temps'])),
        # dcc.Graph(figure=px.scatter(creations[creations.Temps != 0], y="Temps", x="Objet", color="Objet").update_traces(marker_size=10).update_yaxes(tick0=0, dtick=24, gridcolor='black', minor_griddash="dot").update_xaxes(gridcolor='black', minor_griddash="dot")),
        dcc.Graph(figure=px.scatter(creations[creations["Temps"] != 0], y="Temps", x="Objet", color="Objet").update_traces(marker_size=10).update_yaxes(tick0=0, dtick=24, gridcolor='black', minor_griddash="dot").update_xaxes(gridcolor='black', minor_griddash="dot").update_layout(yaxis_title="Temps (h)")),

        html.H2("Activation"),
         dash_table.DataTable(
            columns=[{"name": i, "id": i} for i in ["Date", "Type", "Objet", "Application", "Demandeur", "Temps"]],
            data=activations[["Date", "Type", "Objet", "Application", "Demandeur", "Temps"]].to_dict('records'),
            page_size=10,
            style_data={
                'whiteSpace': 'normal',
                'height': 'auto',
                'backgroundColor': 'white',
                'color': 'black'
            },
            style_header={
                'backgroundColor': 'rgb(30, 30, 30)',
                'color': 'white'
            },
        ),

       dcc.Graph(figure=px.pie(activations, names='Application', color="Application", color_discrete_map=app_colors).update_traces(textinfo='value+label')),
        dcc.Graph(figure=px.histogram(activations, x='Application', color='Objet', color_discrete_map=app_colors, text_auto=True).update_xaxes(categoryorder='total ascending')),
        dcc.Graph(figure=px.histogram(activations, x='Objet', color='Demandeur', text_auto=True).update_xaxes(categoryorder='total ascending')),
        dcc.Graph(figure=px.scatter(activations[activations["Temps"] != 0], y="Temps", x="Objet", color="Objet").update_traces(marker_size=10).update_yaxes(tick0=0, dtick=24, gridcolor='black', minor_griddash="dot").update_xaxes(gridcolor='black', minor_griddash="dot").update_layout(yaxis_title="Temps (h)")),
        html.H2("Erreur"),
          dash_table.DataTable(
            columns=[{"name": i, "id": i} for i in ["Date", "Type", "Objet", "Application", "Demandeur", "Temps"]],
            data=erreurs[["Date", "Type", "Objet", "Application", "Demandeur", "Temps"]].to_dict('records'),
            page_size=10,
            style_data={
                'whiteSpace': 'normal',
                'height': 'auto',
                'backgroundColor': 'white',
                'color': 'black'
            },
            style_header={
                'backgroundColor': 'rgb(30, 30, 30)',
                'color': 'white'
            },
        ),

      dcc.Graph(figure=px.pie(erreurs, names='Application', color="Application", color_discrete_map=app_colors).update_traces(textinfo='value+label')),
        dcc.Graph(figure=px.pie(erreurs, names='Status', color="Status").update_traces(textinfo='value+label')),
        dcc.Graph(figure=px.histogram(erreurs, x='Objet', color='Demandeur', text_auto=True).update_xaxes(categoryorder='total ascending')),
        dcc.Graph(figure=px.histogram(erreurs, x='Objet', color='Traitant', text_auto=True).update_xaxes(categoryorder='total ascending')),
        dcc.Graph(figure=px.scatter(erreurs[erreurs["Temps"] != 0], y="Temps", x="Objet", color="Objet").update_traces(marker_size=10).update_yaxes(tick0=0, dtick=24, gridcolor='black', minor_griddash="dot").update_xaxes(gridcolor='black', minor_griddash="dot").update_layout(yaxis_title="Temps (h)"),
    ),
                  # .update_traces(marker_size=10).update_yaxes(tick0=0, dtick=24, gridcolor='black', minor_griddash="dot").update_xaxes(gridcolor='black', minor_griddash="dot")),
        # dcc.Graph(figure=px.pie(erreurs, names='Traitant').update_xaxes(categoryorder='total ascending')),
        # dcc.Graph(figure=px.histogram(erreurs, x='Application').update_xaxes(categoryorder='total ascending')),

        # html.H2("Annulation"),
        # dash_table.DataTable(
        #     columns=[{"name": i, "id": i} for i in ["Date", "Type", "Objet", "Demandeur", "Traitant"]],
        #     data=annulations[["Date", "Type", "Objet", "Demandeur", "Traitant"]].to_dict('records'),
        #     # page_size=10,
        #     style_data={
        #         'whiteSpace': 'normal',
        #         'height': 'auto',
        #         'backgroundColor': 'white',
        #         'color': 'black'
        #     },
        #     style_header={
        #         'backgroundColor': 'rgb(30, 30, 30)',
        #         'color': 'white'
        #     },
        # ),
        # dcc.Graph(figure=px.pie(annulations, names='Application', color="Application", color_discrete_map=app_colors).update_traces(textinfo='value+label')),
        # dcc.Graph(figure=px.histogram(annulations, x='Objet', color="Application").update_xaxes(categoryorder='total ascending')),
        # dcc.Graph(figure=px.histogram(annulations, x='Objet', color="Demandeur").update_xaxes(categoryorder='total ascending')),
        # dcc.Graph(figure=px.histogram(annulations, x='Objet', color="Traitant").update_xaxes(categoryorder='total ascending')),
        # # dcc.Graph(figure=px.pie(annulations, names='Demandeur').update_traces(textinfo='value+label')),
        # # dcc.Graph(figure=px.pie(annulations, names='Traitant').update_traces(textinfo='value+label')),
        # # dcc.Graph(figure=px.bar(annulations[annulations.Temps != 0]['Temps'])),
        # # dcc.Graph(figure=px.scatter(annulations[annulations.Temps != 0], y="Temps", x="Objet", color="Objet").update_traces(marker_size=10).update_yaxes(tick0=0, dtick=24, gridcolor='black', minor_griddash="dot").update_xaxes(gridcolor='black', minor_griddash="dot")),

        # html.H2("Inforamtion"),
        # dash_table.DataTable(
        #     columns=[{"name": i, "id": i} for i in ["Date", "Type", "Objet", "Application", "Demandeur"]],
        #     data=informations.to_dict('records'),
        #     # page_size=10,
        #     style_data={
        #         'whiteSpace': 'normal',
        #         'height': 'auto',
        #         'backgroundColor': 'white',
        #         'color': 'black'
        #     },
        #     style_header={
        #         'backgroundColor': 'rgb(30, 30, 30)',
        #         'color': 'white'
        #     },
        # ),
        # # dcc.Graph(figure=px.pie(informations, names='Application', color="Application", color_discrete_map=app_colors).update_traces(textinfo='value+label')),
        # dcc.Graph(figure=px.histogram(informations, x='Objet', color='Application', text_auto=True).update_xaxes(categoryorder='total ascending').update_yaxes(dtick=1)),
        # dcc.Graph(figure=px.histogram(informations, x='Objet', color='Demandeur', text_auto=True).update_xaxes(categoryorder='total ascending').update_yaxes(dtick=1)),
        # # dcc.Graph(figure=px.histogram(informations, x='Application').update_xaxes(categoryorder='total ascending')),
        # # dcc.Graph(figure=px.pie(informations, names='Demandeur')),
        # dcc.Graph(figure=px.scatter(informations[informations.Temps != 0], y="Temps", x="Objet", color="Objet").update_traces(marker_size=10).update_yaxes(tick0=0, dtick=24, gridcolor='black', minor_griddash="dot").update_xaxes(gridcolor='black', minor_griddash="dot").update_layout(yaxis_title="Temps (h)")),
    ])
    # def update_graph(col_chosen):
    #     fig = px.histogram(df, x='From')
    #     return fig
    return app

# Run the app
if __name__ == '__main__':
    report().run(debug=True)
