import yaml
import openpyxl
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
import time
import sys
from tqdm import tqdm
YELLOW_HIGHLIGHT = PatternFill(start_color='ffff00',
				   end_color='ffff00',
				   fill_type='solid')
GREEN_HIGHLIGHT = PatternFill(start_color='90ee90',
				   end_color='90ee90',
				   fill_type='solid')
ORANGE_HIGHLIGHT = PatternFill(start_color='ffa500',
				   end_color='ffa500',
				   fill_type='solid')

#Retrieves data from the YAML file and returns a dictionary with the vals.
def get_config_variables():
	with open('config.yaml') as f:
		data = yaml.load(f, Loader=yaml.FullLoader)
		return data

def copy_worksheet(workbook, sheetName):
	source = workbook.active
	target = workbook.copy_worksheet(source)
	target.title = sheetName
	return target

def find_row_with_key(worksheet, key):
	rowNum = 0 #loop through and find row for key
	for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=1, values_only=True):
		rowNum += 1
		for cell in row:
			if(cell != None and isinstance(cell, str)):
				if(key in cell):
				   return rowNum
	return 0

def removeRowsBelow(worksheet):
	rowNum = 0 #loop through and find row for mass flows
	massFlowArr = []
	for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=1, values_only=True):
		rowNum += 1
		for cell in row:
			if(cell == "Mass Flows"):
			   massFlowArr.append(rowNum) #More than one intance of mass flow in sheet
	worksheet.delete_rows(massFlowArr[0] + 1, worksheet.max_row) #Delete from first mass flow cell down

def removeZeroRows(worksheet):
	maxRow = worksheet.max_row
	maxCol = worksheet.max_column
	rowNum = 3
	
	for row in worksheet.iter_rows(min_col=3, max_col=maxCol, min_row=3, max_row=maxRow, values_only=True):
		allString = False
		zeroFlag = False #Assume all zero
		for item in row:
			if(isinstance(item, (float, int))): #Is a number?
				if(item != None):
					if(item != 0):
						zeroFlag = True
		for item in row:
			if (isinstance(item, str)):
				allString = True
		if(zeroFlag == False and allString == False):
			worksheet.delete_rows(rowNum, 1)
		rowNum += 1

def addTitle(worksheet, title):
	worksheet.insert_rows(1)
	worksheet['A1'] = title

def addInOutRows(worksheet):
	worksheet.insert_rows(2)
	worksheet['A2'] = "Section In"
	worksheet.insert_rows(3)
	worksheet['A3'] = "Section Out"
	worksheet.merge_cells('A2:CO2') #WHAT A CRAZY FIX WHY
	worksheet.unmerge_cells('A2:CO2') #Three hours of my life gone

def entropyCalculations(worksheet):
	rowIdx = 0
	maxRow = worksheet.max_row
	maxCol = worksheet.max_column
	flag = False
	for row in worksheet.iter_rows(min_row=1, min_col=1,
		max_col=1, max_row=maxRow, values_only=True):
		rowIdx += 1
		for item in row:
			if(item == "Enthalpy Flow"):
				flag = True #found the one row I needed
				newRow = rowIdx + 1
				worksheet.insert_rows(newRow, 2) #Create new row below Enthalpy
				titleCell = "A" + str(newRow)
				unitCell = "B" + str(newRow)
				worksheet[titleCell] = "Entropy Flow" #Add new name for row
				worksheet[unitCell] = "kW/K"
				titleCell = "A" + str(newRow + 1)
				unitCell = "B" + str(newRow + 1)
				worksheet[titleCell] = "Exergy Flow" #Add new name for row
				worksheet[unitCell] = "MW"
		if(flag):
			break

	conversionFactor = 3600 * 1000

	stream_molar_flow = find_row_with_key(worksheet, "Mole Flows")
	stream_molar_flow_units = "B" + str(stream_molar_flow)
	stream_molar_entropy = find_row_with_key(worksheet, "Molar Entropy")
	stream_molar_entropy_units = "B" + str(stream_molar_entropy)

	for row_cells in worksheet.iter_rows(min_row=newRow, max_row=newRow,
		min_col=3, max_col=maxCol):
		for cell in row_cells:
			firstValue = worksheet[str(cell.column_letter) + str(stream_molar_entropy)]
			secondValue = worksheet[str(cell.column_letter) + str(stream_molar_flow)]
			if(firstValue.value != None and secondValue.value != None):
				if(worksheet[stream_molar_flow_units].value == "kmol/hr" and 
					worksheet[stream_molar_entropy_units].value == "J/kmol-K"):
					calculatedValue = (firstValue.value * secondValue.value)/conversionFactor
				else:
					calculatedValue = 1
					print("Unit error in calculating Entropy Flow, required: \n")
					print("Mole Flow Rate: kmol/hr")
					print("Entropy Mixture: J/kmol-K")
				cell.value = calculatedValue
			#Exergy Calculations
			firstValue = worksheet[str(cell.column_letter) + str(find_row_with_key(worksheet, "Enthalpy Flow"))]
			secondValue = worksheet[str(cell.column_letter) + str(find_row_with_key(worksheet, "Entropy Flow"))]
			if(firstValue.value != None and secondValue.value != None):
				calculatedValue = firstValue.value - (0.3 * secondValue.value)
				cell.offset(row=1).value = calculatedValue

	worksheet.delete_rows(find_row_with_key(worksheet, "Description"), 1)

#Returns column letter of first blank after to and from rows
def find_blank(worksheet):
	to_idx = find_row_with_key(worksheet, "To")
	from_idx = find_row_with_key(worksheet, "From")

	for col in worksheet.iter_cols(min_row=to_idx, max_row=to_idx, min_col=3, max_col=worksheet.max_column):
		for cell in col:
			to_cell = str(cell.column_letter) + str(to_idx)
			from_cell = str(cell.column_letter) + str(from_idx)
			if worksheet[to_cell].value == None and worksheet[from_cell].value == None:
				return cell.column_letter
				#worksheet.delete_cols(cell.column_letter,1)

def removeColumns(worksheet):
	to_idx = find_row_with_key(worksheet, "To")
	from_idx = find_row_with_key(worksheet, "From")
	endCol = column_index_from_string(find_blank(worksheet))
	delArray = []
	for col in worksheet.iter_cols(min_row=to_idx, max_row=to_idx, min_col=3, max_col=endCol):
		for cell in col:
			to_cell = str(cell.column_letter) + str(to_idx)
			from_cell = str(cell.column_letter) + str(from_idx)
			if worksheet[to_cell].value != None and worksheet[from_cell].value != None:
				delArray.append(column_index_from_string(cell.column_letter))
	for i in reversed(delArray):
		worksheet.delete_cols(i,1)
	
def addInOutValues(worksheet):
	to_idx = find_row_with_key(worksheet, "To")
	from_idx = find_row_with_key(worksheet, "From")
	endCol = column_index_from_string(find_blank(worksheet))
	for col in worksheet.iter_cols(min_row=to_idx, max_row=to_idx, min_col=3, max_col=endCol):
		for cell in col:
			write_cell = str(cell.column_letter) + str(1)
			to_cell = str(cell.column_letter) + str(to_idx)
			from_cell = str(cell.column_letter) + str(from_idx)
			if worksheet[to_cell].value != None:
				worksheet[write_cell] = "In"
			if worksheet[from_cell].value != None:
				worksheet[write_cell] = "Out"

	worksheet.delete_cols(endCol, worksheet.max_column) 

def calculate_balance(worksheet):
	enthalpy_flow = find_row_with_key(worksheet, "Enthalpy Flow")
	enthalpy_sum = 0 #Initialize counter
	for col in worksheet.iter_cols(min_row=enthalpy_flow, max_row=enthalpy_flow, min_col=3, max_col=worksheet.max_column):
		lastCol = col
		for cell in col:
			in_out_row = str(cell.column_letter) + "1"
			if worksheet[in_out_row].value == "In": #Check if the first row in this col has In or Out
				enthalpy_sum += cell.value 			#If In, Add to the sum
			if worksheet[in_out_row].value == "Out":
				enthalpy_sum -= cell.value 			#If out, Subtract from sum
	lastCol[0].offset(column=1).value = enthalpy_sum
	lastCol[0].offset(column=1).fill = YELLOW_HIGHLIGHT

	entropy_flow = find_row_with_key(worksheet, "Entropy Flow")
	entropy_sum = 0 #Initialize counter	
	for col in worksheet.iter_cols(min_row=entropy_flow, max_row=entropy_flow, min_col=3, max_col=worksheet.max_column):
		lastCol = col
		for cell in col:
			in_out_row = str(cell.column_letter) + "1"
			if worksheet[in_out_row].value == "In": #Check if the first row in this col has In or Out
				entropy_sum += cell.value 			#If In, Add to the sum
			if worksheet[in_out_row].value == "Out":
				entropy_sum -= cell.value 			#If out, Subtract from sum
	lastCol[0].value = entropy_sum
	lastCol[0].fill = YELLOW_HIGHLIGHT

	exergy_flow = find_row_with_key(worksheet, "Exergy Flow")
	exergy_sum = 0 #Initialize counter	
	for col in worksheet.iter_cols(min_row=exergy_flow, max_row=exergy_flow, min_col=3, max_col=worksheet.max_column):
		lastCol = col
		for cell in col:
			in_out_row = str(cell.column_letter) + "1"
			if worksheet[in_out_row].value == "In": #Check if the first row in this col has In or Out
				exergy_sum += cell.value 			#If In, Add to the sum
			if worksheet[in_out_row].value == "Out":
				exergy_sum -= cell.value 			#If out, Subtract from sum
	lastCol[0].value = exergy_sum
	lastCol[0].fill = YELLOW_HIGHLIGHT
	worksheet[(str((lastCol[0]).column_letter) + "1")].value = "Balances" #Convoluted, just add title to cell

	mass_flows = find_row_with_key(worksheet, "Mass Flows")
	mass_sum = 0 #Initialize counter	
	for col in worksheet.iter_cols(min_row=mass_flows, max_row=mass_flows, min_col=3, max_col=worksheet.max_column):
		lastCol = col
		for cell in col:
			in_out_row = str(cell.column_letter) + "1"
			if worksheet[in_out_row].value == "In": #Check if the first row in this col has In or Out
				mass_sum += cell.value 			#If In, Add to the sum
			if worksheet[in_out_row].value == "Out":
				mass_sum -= cell.value 			#If out, Subtract from sum
	if(abs(mass_sum <= 10)):	#Check if the mass_flow sum is greater than 10, if so MB_Error
		lastCol[0].value = mass_sum
		lastCol[0].fill = YELLOW_HIGHLIGHT
		return 1
	else:	
		print("MB_Error, Mass Flow Sum is: " + str(mass_sum))
		print("See cell: " + str(lastCol[0].coordinate))
		return 0

def step_six(worksheet):
	#Lesson, do not use array for deleting rows because they change dynamically per each deletion
	worksheet.delete_rows(1, 2)
	worksheet.delete_rows(find_row_with_key(worksheet, "Maximum Relative Error"), 1)
	worksheet.delete_rows(find_row_with_key(worksheet, "Cost Flow"), 1)
	worksheet.delete_rows(find_row_with_key(worksheet, "MIXED Substream"), 1)
	worksheet.delete_rows(find_row_with_key(worksheet, "Mass Vapor Fraction"), 1)
	worksheet.delete_rows(find_row_with_key(worksheet, "Mass Liquid Fraction"), 1)
	worksheet.delete_rows(find_row_with_key(worksheet, "Mass Solid Fraction"), 1)
	worksheet.delete_rows(find_row_with_key(worksheet, "Mass Enthalpy"), 1)
	worksheet.delete_rows(find_row_with_key(worksheet, "Mass Entropy"), 1)
	worksheet.delete_rows(find_row_with_key(worksheet, "Mass Density"), 1)
	worksheet.delete_rows(find_row_with_key(worksheet, "Molar Liquid Fraction"), 1)
	worksheet.delete_rows(find_row_with_key(worksheet, "Molar Solid Fraction"), 1)
	removeRowsBelow(worksheet)
	removeZeroRows(worksheet)
	enthalpy_flow = find_row_with_key(worksheet, "Enthalpy Flow")
	enthalpy_flow_units = "B" + str(enthalpy_flow)
	if(worksheet[enthalpy_flow_units].value == "W"):
		print("Enthalpy Flow in Watts, converting to MW")
		for col in worksheet.iter_cols(min_col=3, max_col=worksheet.max_column, 
				min_row=enthalpy_flow, max_row=enthalpy_flow):
			if(col[0].value != None):
				newVal = col[0].value / 1000000
				col[0].value = newVal

def add_text(worksheet, title, overall_text):
	worksheet['A1'] = title
	cnt = 0
	name_array= []
	for col in worksheet.iter_rows(min_row=3, max_row=3):
		for cell in col:
			if(cell.value == "Name"):
				name_array.append(cell.col_idx)
	
	for x in name_array:
		cnt = 0
		for col in worksheet.iter_rows(min_row=64, max_row=110, min_col=x, max_col=x):
			for cell in col:
				if (overall_text[cnt] == "Inlet Stream 1 Name"
					or overall_text[cnt] == "Inlet Stream 2 Name"
					or overall_text[cnt] == "Inlet Stream 3 Name"
					or overall_text[cnt] == "Inlet Stream 4 Name"):
					cell.fill = YELLOW_HIGHLIGHT
				if (overall_text[cnt] == "Outlet Stream 1 Name"
					or overall_text[cnt] == "Outlet Stream 2 Name"
					or overall_text[cnt] == "Outlet Stream 3 Name"
					or overall_text[cnt] == "Outlet Stream 4 Name"
					or overall_text[cnt] == "Outlet Stream 5 Name"
					or overall_text[cnt] == "Outlet Stream 6 Name"):
					cell.fill = ORANGE_HIGHLIGHT
				if (overall_text[cnt] == "Mass balance kg/hr"
					or overall_text[cnt] == "Energy Balance MW"
					or overall_text[cnt] == "Entropy Generation kW/K"):
					cell.fill=GREEN_HIGHLIGHT
				cell.value = overall_text[cnt]
			cnt += 1

	for col in worksheet.iter_rows(min_row=3, max_row=3):
		for cell in col:
			if(cell.value != "Name" and cell.value != None):
				temp = cell.value
				thisCell = cell.offset(row=62)
				thisCell.value = temp

	block_name_array = []
	idx = 0
	for col in worksheet.iter_rows(min_row=2, max_row=2):
		for cell in col:
			if(cell.value != None):
				block_name_array.append(cell.value)
	for col in worksheet.iter_cols(min_row=64, max_row=64):
		for cell in col:
			if cell.value == "Block Type":
				cell.offset(column=1).value = block_name_array[idx]
				idx += 1

def copy_worksheet(workbook, sheetName):
	source = workbook.active
	target = workbook.copy_worksheet(source)
	target.title = sheetName
	return target

def step_seven(worksheet,title):
	addTitle(worksheet,title)
	addInOutRows(worksheet)
	entropyCalculations(worksheet)
	freezeCells = worksheet['C7']
	worksheet.freeze_panes = freezeCells

def step_eight(worksheet):
	freezeCells = worksheet['C7']
	worksheet.freeze_panes = freezeCells
	removeColumns(worksheet)
	addInOutValues(worksheet)

def step_nine(worksheet):
	return calculate_balance(worksheet)

#Looks at the to row, returns array of tuples with format as follows:
#	(to-row-name, stream-name)
def prepare_for_overall_inlet(worksheet):
	return_dict = {}
	to_row = find_row_with_key(worksheet,"To")
	stream_name_row = find_row_with_key(worksheet, "Stream Name") - to_row
	for col in worksheet.iter_cols(min_row=to_row, max_row=to_row,min_col=3):
		for cell in col:
			if cell.value != None:
				if(cell.value in return_dict): #Already in dict 
					return_dict[cell.value].append(cell.offset(row=stream_name_row).value);
				else: #New key
					return_dict[cell.value] = [cell.offset(row=stream_name_row).value]
	return return_dict

def prepare_for_overall_inlet_vals(worksheet):
	return_vals_dict = {}
	to_row = find_row_with_key(worksheet,"To")
	mass_flow_row = find_row_with_key(worksheet,"Mass Flows") - to_row
	enthalpy_flow_row = find_row_with_key(worksheet,"Enthalpy Flow") - to_row
	entropy_flow_row = find_row_with_key(worksheet,"Entropy Flow") - to_row
	for col in worksheet.iter_cols(min_row=to_row, max_row=to_row,min_col=3):
		for cell in col:
			if cell.value != None:
				if(cell.value in return_vals_dict):
					return_vals_dict[cell.value].append((cell.offset(row=mass_flow_row).value, 
						cell.offset(row=enthalpy_flow_row).value,
						cell.offset(row=entropy_flow_row).value))
				else: #New key
					return_vals_dict[cell.value] = [(cell.offset(row=mass_flow_row).value, 
						cell.offset(row=enthalpy_flow_row).value,
						cell.offset(row=entropy_flow_row).value)]
	return return_vals_dict

#Looks at the from row, returns array of tuples with format as follows:
#	(from-row-name, stream-name)
def prepare_for_overall_outlet(worksheet):
	return_dict = {}
	from_row = find_row_with_key(worksheet,"From")
	stream_name_row = find_row_with_key(worksheet, "Stream Name") - from_row
	for col in worksheet.iter_cols(min_row=from_row, max_row=from_row,min_col=3):
		for cell in col:
			if cell.value != None:
				if(cell.value in return_dict): #Already in dict 
					return_dict[cell.value].append(cell.offset(row=stream_name_row).value);
				else: #New key
					return_dict[cell.value] = [cell.offset(row=stream_name_row).value]
	return return_dict

def prepare_for_overall_outlet_vals(worksheet):
	return_vals_dict = {}
	from_row = find_row_with_key(worksheet,"From")
	mass_flow_row = find_row_with_key(worksheet,"Mass Flows") - from_row
	enthalpy_flow_row = find_row_with_key(worksheet,"Enthalpy Flow") - from_row
	entropy_flow_row = find_row_with_key(worksheet,"Entropy Flow") - from_row
	for col in worksheet.iter_cols(min_row=from_row, max_row=from_row,min_col=3):
		for cell in col:
			if cell.value != None:
				if(cell.value in return_vals_dict):
					return_vals_dict[cell.value].append((cell.offset(row=mass_flow_row).value, 
						cell.offset(row=enthalpy_flow_row).value,
						cell.offset(row=entropy_flow_row).value))
				else: #New key
					return_vals_dict[cell.value] = [(cell.offset(row=mass_flow_row).value, 
						cell.offset(row=enthalpy_flow_row).value,
						cell.offset(row=entropy_flow_row).value)]
	return return_vals_dict

#Pretty neat algorithm here if I do say so myself
def step_twelve_inlet(worksheet, inlet_array, inlet_vals_array):
	for col in worksheet.iter_cols(min_row=65, max_row=65,min_col=2):
		for cell in col:
			if cell.value != "Block Name" and cell.value != None:
				for key in inlet_array:
					if cell.value == key:
						arr_len = len(inlet_array[key])
						if(arr_len == 4):
							thisCell = cell.offset(row=13)
							thisCell.value = inlet_array[key][3]
						if(arr_len >= 3):
							thisCell = cell.offset(row=9)
							thisCell.value = inlet_array[key][2]
						if(arr_len >= 2):
							thisCell = cell.offset(row=5)
							thisCell.value = inlet_array[key][1]
						if(arr_len >= 1):
							thisCell = cell.offset(row=1)
							thisCell.value = inlet_array[key][0]
				for key in inlet_vals_array:
					if cell.value == key:
						arr_len = len(inlet_vals_array[key])
						if(arr_len == 4):
							idx = 0
							for val in inlet_vals_array[key][3]:
								thisCell = cell.offset(row=14+idx)
								thisCell.value = val
								idx += 1
						if(arr_len >= 3):
							idx = 0
							for val in inlet_vals_array[key][2]:
								thisCell = cell.offset(row=10+idx)
								thisCell.value = val
								idx += 1
						if(arr_len >= 2):
							idx = 0
							for val in inlet_vals_array[key][1]:
								thisCell = cell.offset(row=6+idx)
								thisCell.value = val
								idx += 1
						if(arr_len >= 1):
							idx = 0
							for val in inlet_vals_array[key][0]:
								thisCell = cell.offset(row=2+idx)
								thisCell.value = val
								idx += 1

def step_twelve_outlet(worksheet, outlet_array, outlet_vals_array):
	for col in worksheet.iter_cols(min_row=65, max_row=65,min_col=2):
		for cell in col:
			if cell.value != "Block Name" and cell.value != None:
				for key in outlet_array:
					if cell.value == key:
						arr_len = len(outlet_array[key])
						if(arr_len == 6):
							thisCell = cell.offset(row=37)
							thisCell.value = outlet_array[key][5]
						if(arr_len >= 5):
							thisCell = cell.offset(row=33)
							thisCell.value = outlet_array[key][4]
						if(arr_len >= 4):
							thisCell = cell.offset(row=29)
							thisCell.value = outlet_array[key][3]
						if(arr_len >= 3):
							thisCell = cell.offset(row=25)
							thisCell.value = outlet_array[key][2]
						if(arr_len >= 2):
							thisCell = cell.offset(row=21)
							thisCell.value = outlet_array[key][1]
						if(arr_len >= 1):
							thisCell = cell.offset(row=17)
							thisCell.value = outlet_array[key][0]
				for key in outlet_vals_array:
					if cell.value == key:
						arr_len = len(outlet_vals_array[key])
						if(arr_len == 6):
							idx = 0
							for val in outlet_vals_array[key][5]:
								thisCell = cell.offset(row=38+idx)
								thisCell.value = val
								idx += 1
						if(arr_len >= 5):
							idx = 0
							for val in outlet_vals_array[key][4]:
								thisCell = cell.offset(row=34+idx)
								thisCell.value = val
								idx += 1
						if(arr_len >= 4):
							idx = 0
							for val in outlet_vals_array[key][3]:
								thisCell = cell.offset(row=30+idx)
								thisCell.value = val
								idx += 1
						if(arr_len >= 3):
							idx = 0
							for val in outlet_vals_array[key][2]:
								thisCell = cell.offset(row=26+idx)
								thisCell.value = val
								idx += 1
						if(arr_len >= 2):
							idx = 0
							for val in outlet_vals_array[key][1]:
								thisCell = cell.offset(row=22+idx)
								thisCell.value = val
								idx += 1
						if(arr_len >= 1):
							idx = 0
							for val in outlet_vals_array[key][0]:
								thisCell = cell.offset(row=18+idx)
								thisCell.value = val
								idx += 1

def get_block_range(worksheet, keyword):
	start_idx = 0
	end_idx = 0
	for col in worksheet.iter_cols(min_row=2, max_row=2):
		for cell in col:
			if cell.value == keyword:
				start_idx = cell.col_idx

	for col in worksheet.iter_cols(min_row=3, max_row=3, min_col=start_idx - 1):
		for cell in col:
			if cell.value == None:
				end_idx = cell.col_idx - 1
				break
		else:
			continue
		break
	return(start_idx, end_idx)

def step_thirteen(worksheet):
	b_range = get_block_range(worksheet, "Heater")
	for col in worksheet.iter_cols(min_col=b_range[0], max_col=b_range[1]):
		for cell in col:
			if cell.value = 
			#Breakpt: How do I know for instance that HeatX is a Heater?

def main():
	inputData = get_config_variables()
	streamWorkbook = inputData["streamBookName"]
	print("Working on: " + str(streamWorkbook))
	wb_stream = openpyxl.load_workbook(streamWorkbook)
	modifiedWS = copy_worksheet(wb_stream, "Aspen Data Tables Modified")
	with tqdm(total=100, file=sys.stdout) as pbar:
		for i in range(1):
			#Begin work on streams workbook
			step_six(modifiedWS)
			pbar.update(25)
			step_seven(modifiedWS, inputData["streamTitle"]) 
			wb_stream.save(streamWorkbook)
			overall = wb_stream.copy_worksheet(modifiedWS)
			overall.title = "Overall"
			pbar.update(25)
			step_eight(overall)
			pbar.update(25)
			check = step_nine(overall)
			wb_stream.save(streamWorkbook)
			if(check == 0):
				sys.exit()
			pbar.update(25)
	
	#Begin work on models workbook 
	inputData = get_config_variables()
	modelWorkbook = inputData["modelBookName"]
	print("Working on:" + str(modelWorkbook))
	wb_block = openpyxl.load_workbook(modelWorkbook)
	overallTitle = inputData["modelTitle"]
	overallWS = copy_worksheet(wb_block, "Overall")
	overall_text_add = inputData["overall_text_add"]
	with tqdm(total=100, file=sys.stdout) as pbar:
		for i in range(1):
			add_text(overallWS, overallTitle, overall_text_add)
			wb_block.save(modelWorkbook)
			pbar.update(25)
			inlet_array = prepare_for_overall_inlet(modifiedWS)
			inlet_vals_array = prepare_for_overall_inlet_vals(modifiedWS)
			step_twelve_inlet(overallWS, inlet_array, inlet_vals_array)
			pbar.update(25)
			outlet_array = prepare_for_overall_outlet(modifiedWS)
			outlet_vals_array = prepare_for_overall_outlet_vals(modifiedWS)
			step_twelve_outlet(overallWS, outlet_array, outlet_vals_array)
			pbar.update(25)
			step_thirteen(overallWS)
			wb_block.save(modelWorkbook)
			pbar.update(25)

if __name__ == '__main__':
	main()

#Small bug: Does the order of the outlets matter (i.e for C302, Column AZ), does it matter that 
#outlet 2 is s40 and not s39? Mine are correct but inverse for >= 	3 outlets.

