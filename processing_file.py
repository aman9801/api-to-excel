import openpyxl
import requests, json

def read_excel(name):
  # data to the workbook
  wb = openpyxl.load_workbook(name) 
  # data from the sheet 1 to a local variable
  sheet1 = wb["Sheet1"]
  # data from the sheet 2 to a local variable
  sheet2 = wb["Sheet2"]
  # data from the sheet 3 to a local variable
  sheet3 = wb["Sheet3"] 
  # getting the value of both the cells
  city1 = sheet1["A1"].value
  city2 = sheet2["A1"].value
  # requesting the data from API
  url1 = f"https://api.openweathermap.org/data/2.5/weather?q={city1}&appid=9821d430bd99ea2cf1e664d59fc7b028"
  url2 = f"https://api.openweathermap.org/data/2.5/weather?q={city2}&appid=9821d430bd99ea2cf1e664d59fc7b028"
  r1 = requests.get(url1)
  r2 = requests.get(url2)
  data1 = r1.json()
  data2 = r2.json()
  data3 = r1.json()
  data4 = r2.json()
  # saving the data for city 1
  sheet1["B2"] = "Current Temperature(Kelvin)"
  sheet1["B3"] = float(data1["main"]["temp"])
  sheet1["C2"] = "Min. Temperature(Kelvin)"
  sheet1["C3"] = float(data1["main"]["temp_min"])
  sheet1["D2"] = "Max. Temperature(Kelvin)"
  sheet1["D3"] = float(data1["main"]["temp_max"])
  sheet1["E2"] = "Longitude"
  sheet1["E3"] = data1["coord"]["lon"]
  sheet1["F2"] = "Latitude"
  sheet1["F3"] = data1["coord"]["lat"]
  sheet1["G2"] = "Humidity %"
  sheet1["G3"] = float(data1["main"]["humidity"])
  sheet1["H2"] = "Description"
  sheet1["H3"] = data1["weather"][0]["description"]
  sheet1["I2"] = "Windspeed(km/h)"
  sheet1["I3"] = data1["wind"]["speed"]
  # saving the data for city 2
  sheet2["B2"] = "Current Temperature(Kelvin)"
  sheet2["B3"] = float(data2["main"]["temp"])
  sheet2["C2"] = "Min. Temperature(Kelvin)"
  sheet2["C3"] = float(data2["main"]["temp_min"])
  sheet2["D2"] = "Max. Temperature(Kelvin)"
  sheet2["D3"] = float(data2["main"]["temp_max"])
  sheet2["E2"] = "Longitude"
  sheet2["E3"] = data2["coord"]["lon"]
  sheet2["F2"] = "Latitude"
  sheet2["F3"] = data2["coord"]["lat"]
  sheet2["G2"] = "Humidity %"
  sheet2["G3"] = float(data2["main"]["humidity"])
  sheet2["H2"] = "Description"
  sheet2["H3"] = data2["weather"][0]["description"]
  sheet2["I2"] = "Windspeed(Km/h)"
  sheet2["I3"] = data2["wind"]["speed"]
  # creating a graph in sheet 1
  chart1 = openpyxl.chart.BarChart()
  chart1.type = "col"
  chart1.style = 10
  chart1.title = f"Temperature in {city1}"
  chart1.y_axis.title = 'Temperature(in Kelvin)'
  chart1.x_axis.title = 'Current - Min - Max'
  data1 = openpyxl.chart.Reference(sheet1, min_col=2, min_row=3, max_col=4, max_row=3)
  titles1 = openpyxl.chart.Reference(sheet1, min_col=2, min_row=2, max_col=4, max_row=2)
  chart1.add_data(data1)
  chart1.set_categories(titles1)
  chart1.shape = 4
  sheet1.add_chart(chart1, 'B6')
  # creating a graph in sheet 2
  chart2 = openpyxl.chart.BarChart()
  chart2.type = "col"
  chart2.style = 10
  chart2.title = f"Temperature in {city2}"
  chart2.y_axis.title = 'Temperature(in Kelvin)'
  chart2.x_axis.title = 'Current - Min - Max'
  data2 = openpyxl.chart.Reference(sheet2, min_col=2, min_row=3, max_col=4, max_row=3)
  titles2 = openpyxl.chart.Reference(sheet2, min_col=2, min_row=2, max_col=4, max_row=2)
  chart2.add_data(data2)
  chart2.set_categories(titles2)
  chart2.shape = 4
  sheet2.add_chart(chart2, 'B6')
  # comparing the data for both the cities
  sheet3['A1'] = "City"
  sheet3['A2'] =  city1
  sheet3['A3'] =  city2
  sheet3["B1"] = "Current Temperature(Kelvin)"
  sheet3["B2"] = float(data3["main"]["temp"])
  sheet3["B3"] = float(data4["main"]["temp"])
  sheet3["C1"] = "Min. Temperature(Kelvin)"
  sheet3["C2"] = float(data3["main"]["temp_min"])
  sheet3["C3"] = float(data4["main"]["temp_min"])
  sheet3["D1"] = "Max. Temperature(Kelvin)"
  sheet3["D2"] = float(data3["main"]["temp_max"])
  sheet3["D3"] = float(data4["main"]["temp_max"])
  sheet3["E1"] = "Longitude"
  sheet3["E2"] = data3["coord"]["lon"]
  sheet3["E3"] = data4["coord"]["lon"]
  sheet3["F1"] = "Latitude"
  sheet3["F2"] = data3["coord"]["lat"]
  sheet3["F3"] = data4["coord"]["lat"]
  sheet3["G1"] = "Humidity %"
  sheet3["G2"] = float(data3["main"]["humidity"])
  sheet3["G3"] = float(data4["main"]["humidity"])
  sheet3["H1"] = "Description"
  sheet3["H2"] = data3["weather"][0]["description"]
  sheet3["H3"] = data4["weather"][0]["description"]
  sheet3["I1"] = "Windspeed(km/h)"
  sheet3["I2"] = data3["wind"]["speed"]
  sheet3["I3"] = data4["wind"]["speed"]
  # creating a graph in sheet 3
  chart3 = openpyxl.chart.BarChart()
  chart3.type = "col"
  chart3.style = 10
  chart3.title = f"Comparison of Windspeed in {city1} and {city2}"
  chart3.y_axis.title = 'Windspeed (in Km/h)'
  data3 = openpyxl.chart.Reference(sheet3, min_col=9, min_row=1, max_row=3)
  titles3 = openpyxl.chart.Reference(sheet3, min_col=1, min_row=2, max_row=3)
  chart3.add_data(data3)
  chart3.set_categories(titles3)
  sheet3.add_chart(chart3, 'B6')

  wb.save(name)
  return True