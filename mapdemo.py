import openpyxl
import openpyxl.utils
import plotly.graph_objects
from us_state_abbrev import us_state_abbrev

def get_excel_rows(file_name):
    excel_file= openpyxl.load_workbook(file_name)
    first_sheet= excel_file.active
    all_data= first_sheet.rows
    return all_data

def main():
    income_data= get_excel_rows("MedianIncomeInflationAdjusted.xlsx")
    state_abbreviation_list = []
    income_changes = []
    for income_row in income_data:
        state_name= income_row[0].value
        if not state_name in us_state_abbrev:
           continue
        income_2018 = income_row[1].value
        income_2008_col= openpyxl.utils.cell.column_index_from_string('z')-1
        income_2008 = income_row[income_2008_col].value
        change_in_income= income_2018 - income_2008
        state_abbreviation = us_state_abbrev[state_name]
        state_abbreviation_list.append(state_abbreviation)
        income_changes.append(change_in_income)

    income_change_map = plotly.graph_objects.Figure(
        data= plotly.graph_objects.Choropleth(
            locations= state_abbreviation_list,
            z= income_changes,
            locationmode="USA-states",
            colorscale= "Electric",
            colorbar_title= "Income Changes since 2008",
        ))
    income_change_map.update_layout(
        title_text = "Inflation Adjusted real Income Changes since 2008",
        geo_scope = "usa")
    income_change_map.show()
main()
