import pandas as pd 
import xlrd as xl 
from pandas import ExcelWriter
from pandas import ExcelFile

def calculate_total_charge(excel_file_name):

	DataF = pd.read_excel(excel_file_name, 'Sheet1')
	time    = DataF["Time (s)"]
	current = DataF["WE(1).Current (A)"]

	Total_Charge = 0

	length = len(time) - 1

	for i in range(length):
		try:
			Total_Charge = Total_Charge + (time[i+1] - time[i]) * current[i]
		except Exception as e:
			print(e, i)
	
	return Total_Charge * -1

	
def eficiencia_faradiaca(DMSO_mass, DMSO_density, DMSO_volume, DMSO_Molecular_Weight, DMSO_Moles, D2O_mass, D2O_density, D2O_volume, Total_volume_in_eppendorf, Reference_mass, Reference_volume, Reference_density, Sample_volume, Sample_density, Dilution_ratio, Total_Volume_NMR, DMSO_concentration_in_NMR_Tube, DMSO_Integral, CH3COOH_Integral, Ratio_CH3COOH_DMSO, CH3COOH_NMR_Concentration, Concentration_ratio, CH3COOH_Concentration, Total_Charge, CH3COOH_Charge, Total_cathodic_volume, Total_CH3COOH_moles, Faradaic_Efficiency, excel_file_name, Reaction_rate, Electroreduction_Time, Total_Bicarbonate_moles, Bicarbonate_Concentration, Conversion, CH3COOH_Concentration_mM):
	
	DMSO_density = 1.1004
	DMSO_volume = (DMSO_mass / DMSO_density) / 1000
	DMSO_Molecular_Weight = 78.13
	DMSO_Moles = DMSO_mass / DMSO_Molecular_Weight
	D2O_density = 1.107
	D2O_volume = (D2O_mass / D2O_density) / 1000
	Total_volume_in_eppendorf = (D2O_volume + DMSO_volume)
	DMSO_concentration_in_ependorf = DMSO_Moles / Total_volume_in_eppendorf
	Reference_density = 1.1
	Reference_volume = (Reference_mass / Reference_density) / 1000
	Sample_volume = 0.0006
	Sample_density = 1
	Total_Volume_NMR = Sample_volume + Reference_volume
	Dilution_ratio = Reference_volume / Total_Volume_NMR 
	DMSO_concentration_in_NMR_Tube = DMSO_concentration_in_ependorf * Dilution_ratio
	Ratio_CH3COOH_DMSO = CH3COOH_Integral / (DMSO_Integral / 2)
	CH3COOH_NMR_Concentration = DMSO_concentration_in_NMR_Tube * Ratio_CH3COOH_DMSO
	Concentration_ratio = Total_Volume_NMR / Sample_volume
	CH3COOH_Concentration = (CH3COOH_NMR_Concentration * Concentration_ratio)
	CH3COOH_Concentration_mM = CH3COOH_Concentration * 1000
	Total_CH3COOH_moles = CH3COOH_Concentration * Total_cathodic_volume
	CH3COOH_Charge = Total_CH3COOH_moles * 6 * 96485.3365
	Faradaic_Efficiency = (CH3COOH_Charge / Total_Charge) * 100
	Reaction_rate = (Total_CH3COOH_moles / Electroreduction_Time) * 1000000000
	Total_Bicarbonate_moles = Bicarbonate_Concentration * Total_cathodic_volume
	Conversion = (Total_CH3COOH_moles / Total_Bicarbonate_moles ) * 100
	

	return DMSO_concentration_in_NMR_Tube, CH3COOH_Concentration, CH3COOH_Concentration_mM, Total_Charge, CH3COOH_Charge, Faradaic_Efficiency, Reaction_rate, Conversion

def main():

	print("Electroreduction performance calculations")
	print("The units of each parameter are specified in parentheses")

	DMSO_mass = float(input("DMSO mass (g): "))
	D2O_mass = float(input("D2O mass (g): "))
	Reference_mass = float(input("Reference mass (g): "))
	DMSO_Integral = float(input("DMSO Integral: "))
	CH3COOH_Integral = float(input("CH3COOH Integral: "))
	Total_cathodic_volume = float(input("Total cathodic volume (L): "))
	Electroreduction_Time = float(input("Electroreduction Time (s): "))
	userimput = input("Excel Name: ")
	Bicarbonate_Concentration = float(input("Bicarbonate concentration (M): "))
	DMSO_density = 1.1004
	DMSO_volume = (DMSO_mass / DMSO_density) / 1000
	DMSO_Molecular_Weight = 78.13
	DMSO_Moles = DMSO_mass / DMSO_Molecular_Weight
	D2O_density = 1.107
	D2O_volume = (D2O_mass / D2O_density) / 1000
	Total_volume_in_eppendorf = (D2O_volume + DMSO_volume)
	DMSO_concentration_in_ependorf = DMSO_Moles / Total_volume_in_eppendorf
	Reference_density = 1.1
	Reference_volume = (Reference_mass / Reference_density) / 1000
	Sample_volume = 0.0006
	Sample_density = 1
	Total_Volume_NMR = Sample_volume + Reference_volume
	Dilution_ratio = Reference_volume / Total_Volume_NMR 
	DMSO_concentration_in_NMR_Tube = DMSO_concentration_in_ependorf * Dilution_ratio
	Ratio_CH3COOH_DMSO = CH3COOH_Integral / (DMSO_Integral / 2)
	CH3COOH_NMR_Concentration = DMSO_concentration_in_NMR_Tube * Ratio_CH3COOH_DMSO
	Concentration_ratio = Total_Volume_NMR / Sample_volume
	CH3COOH_Concentration = (CH3COOH_NMR_Concentration * Concentration_ratio)
	CH3COOH_Concentration_mM = CH3COOH_Concentration * 1000
	excel_file_name = userimput
	Total_Charge = calculate_total_charge(excel_file_name)
	Total_CH3COOH_moles = CH3COOH_Concentration * Total_cathodic_volume
	CH3COOH_Charge = Total_CH3COOH_moles * 6 * 96485.3365
	Faradaic_Efficiency = (CH3COOH_Charge / Total_Charge) * 100
	Reaction_rate = Total_CH3COOH_moles / Electroreduction_Time
	Total_Bicarbonate_moles = Bicarbonate_Concentration * Total_cathodic_volume
	Conversion = (Total_CH3COOH_moles / Total_Bicarbonate_moles ) * 100

	try:
		resultat1, resultat2, resultat3, resultat4, resultat5, resultat6, resultat7, resultat8 = eficiencia_faradiaca (DMSO_mass, DMSO_density, DMSO_volume, DMSO_Molecular_Weight, DMSO_Moles, D2O_mass, D2O_density, D2O_volume, Total_volume_in_eppendorf, Reference_mass, Reference_volume, Reference_density, Sample_volume, Sample_density, Dilution_ratio, Total_Volume_NMR, DMSO_concentration_in_NMR_Tube, DMSO_Integral, CH3COOH_Integral, Ratio_CH3COOH_DMSO, CH3COOH_NMR_Concentration, Concentration_ratio, CH3COOH_Concentration, Total_Charge, CH3COOH_Charge, Total_cathodic_volume, Total_CH3COOH_moles, Faradaic_Efficiency, excel_file_name, Reaction_rate, Electroreduction_Time, Total_Bicarbonate_moles, Bicarbonate_Concentration, Conversion, CH3COOH_Concentration_mM)

		print("DMSO concentration in NMR Tube (M) : ", round(resultat1, 5))
		print("CH3COOH Concentration (M) : ", round(resultat2, 6)) 
		print("CH3COOH Concentration (mM) : ", round(resultat3, 2))
		print("Total Charge : ", round(resultat4, 1))
		print("CH3COOH Charge : ", round(resultat5, 2))
		print("Faradaic Efficiency (%) : ", round(resultat6, 2))
		print("Reaction rate (nM/s) : ", round(resultat7, 2))
		print("Conversion (%) : ", round(resultat8, 4))
		
	
	except Exception as e:
   		print(e)

	
if __name__ == "__main__":
    main()
