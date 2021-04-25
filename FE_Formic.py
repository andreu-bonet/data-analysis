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

	
def eficiencia_faradiaca(DMSO_mass, DMSO_density, DMSO_volume, DMSO_Molecular_Weight, DMSO_Moles, D2O_mass, D2O_density, D2O_volume, Total_volume_in_eppendorf, Reference_mass, Reference_volume, Reference_density, Sample_volume, Sample_density, Dilution_ratio, Total_Volume_NMR, DMSO_concentration_in_NMR_Tube, DMSO_Integral, HCOOH_Integral, Ratio_HCOOH_DMSO, HCOOH_NMR_Concentration, Concentration_ratio, HCOOH_Concentration, Total_Charge, HCOOH_Charge, Total_cathodic_volume, Total_HCOOH_moles, Faradaic_Efficiency, excel_file_name, Reaction_rate, Electroreduction_Time, Total_Bicarbonate_moles, Bicarbonate_Concentration, Conversion, HCOOH_Concentration_mM):
	
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
	Ratio_HCOOH_DMSO = HCOOH_Integral / (DMSO_Integral / 6)
	HCOOH_NMR_Concentration = DMSO_concentration_in_NMR_Tube * Ratio_HCOOH_DMSO
	Concentration_ratio = Total_Volume_NMR / Sample_volume
	HCOOH_Concentration = (HCOOH_NMR_Concentration * Concentration_ratio)
	HCOOH_Concentration_mM = HCOOH_Concentration * 1000
	Total_HCOOH_moles = HCOOH_Concentration * Total_cathodic_volume
	HCOOH_Charge = Total_HCOOH_moles * 2 * 96485.3365
	Faradaic_Efficiency = (HCOOH_Charge / Total_Charge) * 100
	Reaction_rate = (Total_HCOOH_moles / Electroreduction_Time) * 1000000000
	Total_Bicarbonate_moles = Bicarbonate_Concentration * Total_cathodic_volume
	Conversion = (Total_HCOOH_moles / Total_Bicarbonate_moles ) * 100
	

	return DMSO_concentration_in_NMR_Tube, HCOOH_Concentration, HCOOH_Concentration_mM, Total_Charge, HCOOH_Charge, Faradaic_Efficiency, Reaction_rate, Conversion

def main():

	print("Electroreduction performance calculations")
	print("The units of each parameter are specified in parentheses")

	DMSO_mass = float(input("DMSO mass (g): "))
	D2O_mass = float(input("D2O mass (g): "))
	Reference_mass = float(input("Reference mass (g): "))
	DMSO_Integral = float(input("DMSO Integral: "))
	HCOOH_Integral = float(input("HCOOH Integral: "))
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
	Ratio_HCOOH_DMSO = HCOOH_Integral / (DMSO_Integral / 6)
	HCOOH_NMR_Concentration = DMSO_concentration_in_NMR_Tube * Ratio_HCOOH_DMSO
	Concentration_ratio = Total_Volume_NMR / Sample_volume
	HCOOH_Concentration = (HCOOH_NMR_Concentration * Concentration_ratio)
	HCOOH_Concentration_mM = HCOOH_Concentration * 1000
	excel_file_name = userimput
	Total_Charge = calculate_total_charge(excel_file_name)
	Total_HCOOH_moles = HCOOH_Concentration * Total_cathodic_volume
	HCOOH_Charge = Total_HCOOH_moles * 2 * 96485.3365
	Faradaic_Efficiency = (HCOOH_Charge / Total_Charge) * 100
	Reaction_rate = Total_HCOOH_moles / Electroreduction_Time
	Total_Bicarbonate_moles = Bicarbonate_Concentration * Total_cathodic_volume
	Conversion = (Total_HCOOH_moles / Total_Bicarbonate_moles ) * 100

	try:
		resultat1, resultat2, resultat3, resultat4, resultat5, resultat6, resultat7, resultat8 = eficiencia_faradiaca (DMSO_mass, DMSO_density, DMSO_volume, DMSO_Molecular_Weight, DMSO_Moles, D2O_mass, D2O_density, D2O_volume, Total_volume_in_eppendorf, Reference_mass, Reference_volume, Reference_density, Sample_volume, Sample_density, Dilution_ratio, Total_Volume_NMR, DMSO_concentration_in_NMR_Tube, DMSO_Integral, HCOOH_Integral, Ratio_HCOOH_DMSO, HCOOH_NMR_Concentration, Concentration_ratio, HCOOH_Concentration, Total_Charge, HCOOH_Charge, Total_cathodic_volume, Total_HCOOH_moles, Faradaic_Efficiency, excel_file_name, Reaction_rate, Electroreduction_Time, Total_Bicarbonate_moles, Bicarbonate_Concentration, Conversion, HCOOH_Concentration_mM)

		print("DMSO concentration in NMR Tube (M) : ", round(resultat1, 5))
		print("HCOOH Concentration (M) : ", round(resultat2, 6)) 
		print("HCOOH Concentration (mM) : ", round(resultat3, 2))
		print("Total Charge : ", round(resultat4, 1))
		print("HCOOH Charge : ", round(resultat5, 2))
		print("Faradaic Efficiency (%) : ", round(resultat6, 2))
		print("Reaction rate (nM/s) : ", round(resultat7, 2))
		print("Conversion (%) : ", round(resultat8, 4))
		
	
	except Exception as e:
   		print(e)

	
if __name__ == "__main__":
    main()
