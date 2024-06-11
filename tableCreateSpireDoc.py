from spire.doc import *
from spire.doc.common import *

def generate_spire_document(input_dict_hotel, input_dict_car):
    # Create a Document object
    doc = Document()
    doc.LoadFromFile('mergeDocs/table_page.docx')
    # Add a section
    section = doc.Sections[0]

    # Create a table
    hotel_table = Table(doc, True)

    # Set the width of table
    hotel_table.PreferredWidth = PreferredWidth(WidthType.Percentage, int(100))

    # Set the border of table
    hotel_table.TableFormat.Borders.BorderType = BorderStyle.Single
    hotel_table.TableFormat.Borders.Color = Color.get_Black()

    # Define table data for the first table (Hotel) with column headings
    hotel_table_data = [["Destination", "Hotel", "Price per Night"]]

    # Add data to hotel_table_data
    for city, details in input_dict_hotel.items():
        hotel, price = details
        hotel_table_data.append([city, hotel, price])

    # Add rows to the first table (Hotel)
    for rowData in hotel_table_data:
        row = hotel_table.AddRow(False, len(rowData))
        row.Height = 20.0
        for i, col in enumerate(rowData):
            cell = row.Cells[i]
            cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
            paragraph = cell.AddParagraph()
            paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
            paragraph.AppendText(col)

    # Add the hotel table to the section
    section.Tables.Add(hotel_table)

    # Create the second table (Car)
    car_table = Table(doc, True)

    # Set the width of table
    car_table.PreferredWidth = PreferredWidth(WidthType.Percentage, 100)

    # Set the border of table
    car_table.TableFormat.Borders.BorderType = BorderStyle.Single
    car_table.TableFormat.Borders.Color = Color.get_Black()

    # Define table data for the second table (Car) with column headings
    car_table_data = [["Destination", "Car", "Fare"]]

    # Add data to car_table_data
    for city, details in input_dict_car.items():
        car, fare = details
        car_table_data.append([city, car, fare])

    # Add rows to the second table (Car)
    for rowData in car_table_data:
        row = car_table.AddRow(False, len(rowData))
        row.Height = 20.0
        for i, col in enumerate(rowData):
            cell = row.Cells[i]
            cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
            paragraph = cell.AddParagraph()
            paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
            paragraph.AppendText(col)

    # Add the car table to the section
    section.Tables.Add(car_table)

    # Save the document
    doc.SaveToFile("tableSpire.docx", FileFormat.Docx2019)
    doc.Close()

# # Example usage:
# hotel_data = {
#     "New York": ("Hilton", "$200"),
#     "London": ("Marriott", "$150"),
#     "Paris": ("Sheraton", "$180")
# }

# car_data = {
#     "New York": ("Sedan", "$50"),
#     "London": ("SUV", "$70"),
#     "Paris": ("Convertible", "$90")
# }

# generate_spire_document(hotel_data, car_data)