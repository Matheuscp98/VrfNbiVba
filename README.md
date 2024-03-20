Description and Use of the Code: 
<br>
<br>
	This code aims to perform the complete VRF-NBI process outlined in the article "A Hybrid Multivariate Normal Boundary Intersection Approach with Post-Optimization Assisted by Mixture Design of Experiments" (self-authored code), starting from individual 		optimization to the varimax rotated factor-normal boundary intersection. Additionally, it includes formatting functions <br>
	The purpose of this code is to demonstrate how everything can be run in Excel using VBA programming <br>
<br>
<br>
*Worksheets used:<br>
	-Data<br>
	-VRF-NBI<br>
	-NBI-8Y<br>
	-Post-Optimization (Mixture)<br>
	-Post-Optimization (RSM)<br>
	-Generalized Distance and Entropy<br>
	-Ellipse (Data)<br>
	-Ellipse (Plots)<br>
<br>
<br>
*Routines present in VBA:<br>
	-Functions for Individual Optimization<br>
		-PayoffMatrix3: Calculates the Payoff Matrix with the original data, then copies and pastes these values for the NBI process, storing them in the same worksheet<br>
		-PayoffMatrix8: Calculates the Payoff Matrix with the rotated factors, then copies and pastes these values for the NBI process, storing them in the same worksheet<br>
	-Functions for VRF-NBI<br>
		-NBIASolve: Execution of the VRF-NBI method with rotated factors, it uses the previous points in each iteration, thus, they are kept and usable for all 66 iterations<br>
		-NBIOSolve: Execution of the VRF-NBI method with rotated factors, it uses the average optimal points (obtained from NBIASolve) in each iteration.<br>
		-NBIZSolve: Execution of the VRF-NBI method with rotated factors, it uses zeroed points in each iteration, thus, the points will be zeroed for each of the 66 iterations<br>
	-Functions for NBI-8Y<br>
		-NBIA8Solve: Execution of the classical NBI method with original variables, it uses the previous points in each iteration, thus, they are kept and usable for all 792 iterations<br>
		-NBIO8Solve: Execution of the classical NBI method with original variables, it uses the average optimal points (obtained from NBIASolve) in each iteration<br>
		-NBIZ8Solve: Execution of the classical NBI method with original variables, it uses zeroed points in each iteration, thus, the points will be zeroed for each of the 792 iterations<br>
	-Functions for NBI post-optimization<br>
		-OptiIndPost: Executes individual optimization of the metrics addressed in the post-optimization<br>
		-NBIPostRSM: Executes post-optimization with RSM of the ideal weights found in the post-optimization to identify the ideal point<br>
	-Additional Functions<br>
		-EnableFullScreen: Enable full-screen mode<br>
		-DisableFullScreen: Disable full-screen mode<br>
		-SaveWorkbook: Save workbook<br>
		-ClearCells3: Clear cells for VRF-NBI<br>
		-ClearCells8: Clear cells for NBI-8Y<br>
		-ClearCellsPost: Clear cells for post-optimization<br>
		-SearchPoints3: Search points for VRF-NBI<br>
		-SearchPoints8: Search points for NBI-8Y<br>
		-SavePoints3: Save points for VRF-NBI<br>
		-SavePoints8: Save points for NBI-8Y<br>
<br>
<br>
*Contact:
-E-mail: matheusc_pereira@hotmail.com
-Linkedin: https://www.linkedin.com/in/matheuscostapereira/
-Lattes: https://lattes.cnpq.br/7025666927284220
