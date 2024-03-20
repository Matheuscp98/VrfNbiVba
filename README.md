Description and Use of the Code:
	This code aims to perform the complete VRF-NBI process outlined in the article "A Hybrid Multivariate Normal Boundary Intersection Approach with Post-Optimization Assisted by Mixture Design of Experiments" (self-authored code), starting from individual optimization to the varimax rotated factor-normal boundary intersection. Additionally, it includes formatting functions
	The purpose of this code is to demonstrate how everything can be run in Excel using VBA programming

Worksheets used:
	Data
	VRF-NBI
	NBI-8Y
	Post-Optimization (Mixture)
	Post-Optimization (RSM)
	Generalized Distance and Entropy
	Ellipse (Data)
	Ellipse (Plots)

Routines present in VBA:
	Functions for Individual Optimization
		PayoffMatrix3: Calculates the Payoff Matrix with the original data, then copies and pastes these values for the NBI process, storing them in the same worksheet
		PayoffMatrix8: Calculates the Payoff Matrix with the rotated factors, then copies and pastes these values for the NBI process, storing them in the same worksheet
	Functions for VRF-NBI
		NBIASolve: Execution of the VRF-NBI method with rotated factors, it uses the previous points in each iteration, thus, they are kept and usable for all 66 iterations
		NBIOSolve: Execution of the VRF-NBI method with rotated factors, it uses the average optimal points (obtained from NBIASolve) in each iteration.
		NBIZSolve: Execution of the VRF-NBI method with rotated factors, it uses zeroed points in each iteration, thus, the points will be zeroed for each of the 66 iterations
	Functions for NBI-8Y
		NBIA8Solve: Execution of the classical NBI method with original variables, it uses the previous points in each iteration, thus, they are kept and usable for all 792 iterations
		NBIO8Solve: Execution of the classical NBI method with original variables, it uses the average optimal points (obtained from NBIASolve) in each iteration
		NBIZ8Solve: Execution of the classical NBI method with original variables, it uses zeroed points in each iteration, thus, the points will be zeroed for each of the 792 iterations
	Functions for NBI post-optimization
		OptiIndPost: Executes individual optimization of the metrics addressed in the post-optimization
		NBIPostRSM: Executes post-optimization with RSM of the ideal weights found in the post-optimization to identify the ideal point
	Additional Functions
		EnableFullScreen: Enable full-screen mode
		DisableFullScreen: Disable full-screen mode
		SaveWorkbook: Save workbook
		ClearCells3: Clear cells for VRF-NBI
		ClearCells8: Clear cells for NBI-8Y
		ClearCellsPost: Clear cells for post-optimization
		SearchPoints3: Search points for VRF-NBI
		SearchPoints8: Search points for NBI-8Y
		SavePoints3: Save points for VRF-NBI
		SavePoints8: Save points for NBI-8Y
	
Contact:
E-mail: matheusc_pereira@hotmail.com
Linkedin: https://www.linkedin.com/in/matheuscostapereira/
Lattes: https://lattes.cnpq.br/7025666927284220
