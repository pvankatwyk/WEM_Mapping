import pandas as pd
import numpy as np

def productionUpdate():
    data = pd.read_excel(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\WEM Financial\Oil production\WEM Uintah Well Production.xlsx', sheet_name = 'Well Production Data')
    data.columns = ['well_name', 'pruno', 'api', 'wi', 'graph_name', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10','11',
                    '12', '13', '14', '15', '16', '17', '18', '19', '20','21', '22', '23', '24', '25', '26', '27', '28', '29', '30', 'total']
    data = data.dropna(subset = ['graph_name'])

    entity_number = np.array(data.pruno).astype(int)
    well_list = np.array(data.well_name).astype(str)

    data_dict = {}
    for idx, value in enumerate(entity_number):
        path = r'https://oilgas.ogm.utah.gov/oilgasweb/live-data-search/lds-disp/disp-grid.xhtml?pruno=' + str(value)
        data = pd.read_html(path)
        df = data[1]
        key = well_list[idx]
        data_dict[key] = df['Oil']['Produced']

    out = pd.DataFrame(data_dict)
    out = out.transpose()
    fp = r'C:/Users/Accounting/Downloads/production.csv'
    out.to_csv(r'C:/Users/Accounting/Downloads/production.csv')
    out_text = 'Done, see ' + fp + '.'
    return out_text