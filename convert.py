import comtypes.client
import os

def ppt_to_pdf(input_path, output_path):
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
    presentation.SaveAs(output_path, 32) 
    presentation.Close()
    powerpoint.Quit()

if __name__ == "__main__":
    input_file = os.path.abspath("Bike Ride Forecasting.pptx")
    output_file = os.path.abspath("output.pdf")

    ppt_to_pdf(input_file, output_file)

    print("Conversion successful!")
    os.startfile(output_file)
