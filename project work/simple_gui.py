
from pathlib import Path
from docxtpl import DocxTemplate

document_path = Path(__file__).parent / "example.docx"
doc = DocxTemplate(document_path)

layout = [
    [sg.Text("Date"), sg.Input(key="DATE")],
    [sg.Text("Name"), sg.Input(key="NAME")],
    #[sg.Text("Staff ID"), sg.Input(key="STAFFID")],
    #[sg.Text("Father Name"), sg.Input(key="FATHERNAME")],
    #[sg.Text("Address"), sg.Input(key="ADDRESS")],
    #[sg.Text("City"), sg.Input(key="CITY")],
    #[sg.Text("Distt/State"), sg.Input(key="DISTT_STATE")],
    #[sg.Text("Designation"), sg.Input(key="DESIGNATION")],
    #[sg.Text("Department"), sg.Input(key="DEPARTMENT")],
    #[sg.Text("College"), sg.Input(key="COLLEGE")],
    #[sg.Text("From"), sg.Input(key="FROM")],
    #[sg.Text("To"), sg.Input(key="TO")],
    #[sg.Text("Resignation Date"), sg.Input(key="RESIGNATION_DATE")],
    [sg.Button("Save"), sg.Exit()],
]

window = sg.Window("Experience Certificate", layout,element_justification="right")

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == "Save":
        #print(event, values)
        doc.render(values)
        output_path = Path(__file__).parent / f"{values['NAME']}.docx"
        doc.save(output_path)
        sg.popup("file saved",f"file has been saved here:{output_path}")
window.close()