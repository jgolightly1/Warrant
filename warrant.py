import sys
import FreeSimpleGUI as sg
from pathlib import Path
from docxtpl import DocxTemplate
import datetime

today=datetime.datetime.today()

document_path = Path(__file__).parent / "Template.docx"
doc = DocxTemplate(document_path)

# ----------- Set the Theme -----------
sg.theme('DarkGrey5')

# ----------- Set Certain Keys for DOCXTPL -----------
context = {
    'RESIDENTIAL':'RESIDENTIAL', 'CSAM':'CSAM', 'GANG':'GANG', 'NARCOTICS':'NARCOTICS', 'MURDER':'MURDER', 'BAC':'BAC',
    'ATT':'ATT', 'VERIZON':'VERIZON', 'INFO':'INFO'
    }

# ----------- Create Different Columns for Different Warrants -----------
col1 = sg.Column([[sg.Frame('Case Identifiers:',
                            [[sg.Column([[sg.Text('What would you like to name this file?'), sg.Input(key='REF')],
                            [sg.Text('What is your assigned Case Number?'), sg.Input(key='CASE')],
                            ])]])]])

col2 = sg.Column([[sg.Frame('Information and Background:', [[sg.Text(), sg.Column([
    [sg.Text('What County are you in?')],
    [sg.Input(key='COUNTY')],
    [sg.Text('What is your Name?')],
    [sg.Input(key='NAME')],
    [sg.Text('Who do you work for?')],
    [sg.Input(key='AGENCY')],
    [sg.Text('What is your Rank?')],
    [sg.Input(key='RANK')],
    [sg.Text('How long have you been a police officer (in years)?')],
    [sg.Input(key='YEARS')],
    [sg.Text('What crime are you investigating?')],
    [sg.Input(key='CRIME')],
    [sg.Text("What's the statute number?")],
    [sg.Input(key='RSMO')],
    [sg.Text('What day (1,2,3) did this occur?')],
    [sg.Input(key='DAY')],
    [sg.Text('What Month?')],
    [sg.Input(key='MONTH')],
    [sg.Text('What Year?')],
    [sg.Input(key='YEAR')],
    ])]])]])

col3 = sg.Column([
    [sg.Frame('Categories:',[[sg.Push(), sg.Text('Does your warrant require any special language?'), sg.Push()],
                            [sg.Checkbox('CSAM', key='CSAM'), sg.Checkbox('GANG', key='GANG'), sg.Checkbox('NARCOTICS', key='NARCOTICS'),
                             sg.Checkbox('MURDER', key='MURDER')]])]])

col4 = sg.Column([
    [sg.Frame('Information:', [[sg.Text(), sg.Column([
    [sg.Radio('Residential', 'radio1', key='RESIDENTIAL', default=True)],
    [sg.Text('Street Address:')],
    [sg.Input(key='STREET')],
    [sg.Text('City:')],
    [sg.Input(key='CITY')],
    [sg.Text('State:')],
    [sg.Input(key='STATE')],
    [sg.Text('Zip Code:')],
    [sg.Input(key='ZIP')],
    [sg.Text('Finish the sentence: The residence is described as a')],
    [sg.Multiline(key='DESCRIPTION', size=(43,10))]])]])]])

col5 = sg.Column([
    [sg.Frame('Device Information:', [[sg.Text(), sg.Column([
    [sg.Radio('Cell Phone', 'radio1', key='CELL_PHONE', default=True)],
    [sg.Text('Brand:')],
    [sg.Input(key='BRAND')],
    [sg.Text('Model:')],
    [sg.Input(key='MODEL')],
    [sg.Text('Operating System (Android or iOS):')],
    [sg.Input(key='OS')],
    [sg.Text('Color:')],
    [sg.Input(key='PHONE_COLOR')],
    [sg.Text('IMEI or other Identifying Number')],
    [sg.Input(key='IMEI')],
    [sg.Text('Evidence Item Number')],
    [sg.Input(key='ITEM_NUMBER')]])]])]])

col6 = sg.Column([
    [sg.Frame('Vehicle Information:', [[sg.Text(), sg.Column([
    [sg.Radio('Vehicle', 'radio1', key='VEHICLE', default=True)],
    [sg.Text('Year:')],
    [sg.Input(key='VEH_YEAR')],
    [sg.Text('Make:')],
    [sg.Input(key='VEH_MAKE')],
    [sg.Text('Model:')],
    [sg.Input(key='VEH_MODEL')],
    [sg.Text('License Plate Number:')],
    [sg.Input(key='LIC_PLATE')],
    [sg.Text('Vehicle Identification Number')],
    [sg.Input(key='VIN')],
    [sg.Text('Color')],
    [sg.Input(key='VEH_COLOR')]])]])]])

col7 = sg.Column([
    [sg.Frame('Suspect Information:', [[sg.Text(), sg.Column([
    [sg.Radio('DNA', 'radio1', key='DNA', default=True)],
    [sg.Text("What is the suspect's name?")],
    [sg.Input(key='SUS_NAME')],
    [sg.Text("What is the suspect's DOB?")],
    [sg.Input(key='SUS_DOB')],
    [sg.Text("What is the suspect's SOC?")],
    [sg.Input(key='SUS_SOC')],
    [sg.Text("What's the address where the suspect is currently located?")],
    [sg.Input(key='SUS_ADDRESS')],
    [sg.Text("City?")],
    [sg.Input(key='SUS_CITY')],
    [sg.Text('State?')],
    [sg.Input(key='SUS_STATE')]])]])]])

col8 = sg.Column([
    [sg.Frame('Vehicle Information:', [[sg.Text(), sg.Column([
    [sg.Radio('Infotainment', 'radio1', key='INFO', default=True)],
    [sg.Text('Year:')],
    [sg.Input(key='VEH_YEAR')],
    [sg.Text('Make:')],
    [sg.Input(key='VEH_MAKE')],
    [sg.Text('Model:')],
    [sg.Input(key='VEH_MODEL')],
    [sg.Text('License Plate Number:')],
    [sg.Input(key='LIC_PLATE')],
    [sg.Text('Vehicle Identification Number')],
    [sg.Input(key='VIN')],
    [sg.Text('Color')],
    [sg.Input(key='VEH_COLOR')]])]])]])

col0 = sg.Column([[sg.Frame('Actions:',
                            [[sg.Column([[sg.Button('Save'), sg.Exit()]])]])]])

# ----------- Create Different Windows -----------
def make_welcome():        
    layout = [ 
    [sg.Text('Sikeston Department of Public Safety', font=('helvetica', '14'), pad=(100,10))],
    [sg.Text('Criminal Investigations Unit', font=('helvetica', '13'))],
    [sg.Image('Patch.png', pad=(50,10))],
    [sg.Text('General Warrants', font=('helvetica', '12'))],
    [sg.Button('Residential', font=('helvetica', '12'), pad=(10,10)), 
     sg.Button('Vehicle', font=('helvetica', '12'), pad=(10,10)),
     sg.Button('Vehicle Infotainment', font=('helvetica', '12'), pad=(10,10)),
     sg.Button('Cell Phone', font=('helvetica', '12'), pad=(10,10)),
     sg.Button('DNA', font=('helvetica', '12'), pad=(10,10))],
    [sg.Text('Search Warrant Creation Tool', font=('helvetica', '12'))],
    [sg.Text('Version: 1.0', font=('helvetica', '12'))],
    [sg.Button('Exit', font=('helvetica', '12'), pad=(10,10))]]
    
    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, element_justification='c', finalize=True)

def make_residential():
    layout = [sg.Push(), col1, sg.Push()], [col2, col4], [sg.Push(), [sg.Push(), col3, sg.Push()], col0, sg.Push()]
    
    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, font=('helvetica', '12'), element_padding=(5,5), finalize=True)

def make_cellphone():
    layout = [sg.Push(), col1, sg.Push()], [col2, col5], [sg.Push(), col0, sg.Push()]
    
    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, font=('helvetica', '12'), element_padding=(5,5), finalize=True)

def make_vehicle():
    layout = [sg.Push(), col1, sg.Push()], [col2, col6], [sg.Push(), [sg.Push(), col3, sg.Push()], col0, sg.Push()]
    
    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, font=('helvetica', '12'), element_padding=(5,5), finalize=True)

def make_DNA():
    layout = [sg.Push(), col1, sg.Push()], [col2, col7], [sg.Push(), col0, sg.Push()]
    
    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, font=('helvetica', '12'), element_padding=(5,5), finalize=True)

def make_infotainment():
    layout = [sg.Push(), col1, sg.Push()], [col2, col8], [sg.Push(), col0, sg.Push()]
    
    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, font=('helvetica', '12'), element_padding=(5,5), finalize=True)

window1, window2, window3 = make_welcome(), None, None

# ----------- Event Loop -----------
while True:
    window, event, context = sg.read_all_windows()
    if window == window2 and event in (sg.WIN_CLOSED, 'Exit'):
        break

    if window == window1:
        if event == 'Residential':
            window1.close()
            window2 = make_residential()
                
        if event == 'Cell Phone':
            window1.close()
            window2 = make_cellphone()

        if event == 'Vehicle':
            window1.close()
            window2 = make_vehicle()

        if event == 'DNA':
            window1.close()
            window2 = make_DNA()            

        if event == 'Vehicle Infotainment':
            window1.close()
            window2 = make_infotainment() 

    if event == "Save":
        # Add calculated fields to our dict

        context["TODAY"] = today.strftime("%m/%d/%Y")

        # Render the template, save new word document & inform user
        doc.render(context)

        output_path = Path(__file__).parent / 'Complete/' / f"{context['REF']}-Search Warrant.docx"
        doc.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")
        break

    if event == 'Exit':
        break

window.close()
