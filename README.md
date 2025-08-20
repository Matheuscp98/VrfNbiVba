# 📊 VRF-NBI: VBA Code

## 📝 Description

This repository contains **VBA code** that implements the **Varimax Rotated Factor Normal Boundary Intersection (VRF-NBI)** method.

The code is designed to run **inside Excel** and covers everything from **individual optimization** to the full **Normal Boundary Intersection (NBI)** process, including **post-optimization routines**, **utilities for formatting**, and **visualization**.  

> For **Python implementation** of the method, see: [**VRF-NBI in Python**](https://github.com/Matheuscp98/VRF_NBI).

---

## 📚 Publications

In addition to ongoing manuscripts, the following public papers are already available:

- **Nonlinear Multiobjective Optimization of Efficiency Conditions using a CFD-DOE Hybrid Approach: A Practical Application in Centrifugal Fans for Industrial Ovens**  
  [Read here](https://www.sciencedirect.com/science/article/pii/S2451904925006900)
- **Strategies in Decision Making in a Multiobjective Context: Integration of DOE, NBI, and CFD in the Optimization of a Centrifugal Fan**  
  [Read here](https://publicacoes.softaliza.com.br/cilamce/article/view/10211/7235)
- **Optimizing Mesh and CFD Simulation Performance: A Multivariate Analysis Approach**  
  [Read here](https://publicacoes.softaliza.com.br/cilamce/article/view/8110/6998)
- **Multiobjective optimization of CFD simulation costs and quality: Exploring the Pareto frontier** (*Otimização Multiobjetivo de Custos e Qualidade de Simulações de CFD: Explorando a Fronteira de Pareto*)  
  [Read here](https://proceedings.science/sbpo/sbpo-2024/trabalhos/otimizacao-multiobjetivo-de-custos-e-qualidade-de-simulacoes-de-cfd-explorando-a?lang=pt-br)

---

## 🛠️ How to Use

1. **Clone or download** this repository to your **local machine**.  
2. **Open** the **Excel file** containing the **VBA macros**.  
3. **Enable macros** in **Excel**.  
4. **Run** the desired **VBA routines** from the **VBA editor** or assign them to **buttons** in your **Excel sheets**.  
5. Check the **detailed code comments** and **visual figures** for easier understanding!  

---

## 📁 Worksheets Included

| Worksheet             | Description                                                                   |
|-----------------------|-------------------------------------------------------------------------------|
| **NBI-VRF 3x3Y**      | Multiobjective optimization (MO) for 3 inputs and 3 outputs.                      |
| **NBI-VRF Post**      | Post-optimization using VRF-NBI for Mixture Design, with metrics as outputs.  |
| **NBI-VRF RSM**       | Multiobjective optimization (MO) to find the "optimal of the optimals".           |
| **NBI-VRF 3x8Y**      | Multiobjective optimization (MO) for 8 inputs and 8 outputs.                      |

---

## 🖼️ Figures

| Example                                      | Screenshot                        | Description                                        |
|-----------------------------------------------|-----------------------------------|----------------------------------------------------|
| NBI-VRF - Buttons                            | ![NBI-VRF - Buttons](NBI-VRF_Buttons.jpg) | Existing buttons in the NBI-VRF sheet.             |
| NBI-VRF 3x3Y - Home                          | ![NBI-VRF 3x3Y - Home](NBI-VRF-3x3Y_Home.jpg) | Initial screen of the NBI-VRF 3x3Y sheet.          |
| NBI-VRF 3x3Y - Metrics, Constraints, VarCovar| ![NBI-VRF-3x3Y_MetricsConstrainsandVarCovar](NBI-VRF-3x3Y_MetricsConstrainsandVarCovar.jpg) | Metrics, constraints, payoff matrix, var-covar.    |
| NBI-VRF 3x3Y - NBI Points                    | ![NBI-VRF 3x3Y - NBI Points](NBI-VRF3x3Y_NBIPoints.jpg) | Optimal points by Zeroed NBI method.               |
| NBI-VRF 3x3Y - Search Points                 | ![NBI-VRF 3x3Y - Search Points](NBI-VRF3x3Y_SearchPoints.jpg) | Search using 3 VRF-NBI methods.                    |
| NBI-VRF 3x8Y - Home                          | ![NBI-VRF 3x8Y - Home](NBI-VRF-3x8Y_Home.jpg) | Initial screen of the NBI-VRF 3x8Y sheet.          |
| NBI-VRF Post - Home                          | ![NBI-VRF Post - Home](NBI-VRF-Post_Home.jpg) | Initial screen of the NBI-VRF Post sheet.          |
| NBI-VRF RSM - Home                           | ![NBI-VRF-RSM_Home](NBI-VRF-RSM_Home.jpg) | Initial screen of the NBI-VRF RSM sheet.           |

---

## ⚙️ Main VBA Routines

<details>
<summary><b>Variable Creation</b></summary>

- `Variables`: Declares and initializes necessary variables.
</details>

<details>
<summary><b>Individual Optimization</b></summary>

- `IndividualOptimization`: Executes individual optimization routines.
- `OptimizationTable`: Generates and manages the optimization table.
</details>

<details>
<summary><b>VRF-NBI Routines</b></summary>

- `ZeroedNBI`: NBI process with zeroed initial points.
- `PreviousNBI`: NBI process using previous optimal points.
- `OptimalNBI`: Finds optimal points.
- `ZeroedPostNBI`: Post-optimization with zeroed points.
- `PreviousPostNBI`: Post-optimization using previous optimal points.
- `NBIPostRSM`: Post-optimization using RSM.
</details>

<details>
<summary><b>Additional Utilities</b></summary>

- `SearchPoints`: Searches for points.
- `Clear`: Clears designated cells.
- `EnableFullScreen` / `DisableFullScreen`: Full-screen controls.
- `Save`: Saves the workbook.
- `SavePoint`: Saves optimization points.
- `ClearPoints`: Clears saved points.
- `Home`: Go to the main worksheet.
</details>

---

## 📬 Contact

<a href="mailto:matheusc_pereira@hotmail.com">
  <img src="https://i.ibb.co/k6Ddn36k/email.png" alt="E-mail" height="60"/>
</a>
<a href="https://www.linkedin.com/in/matheuscostapereira/">
  <img src="https://i.ibb.co/Kx4rZxdr/linkedin.png" alt="LinkedIn" height="60"/>
</a>
<a href="https://scholar.google.com.br/citations?user=1iDBIzYAAAAJ&hl=en-us">
  <img src="https://i.ibb.co/SwsRKK1t/scholar.png" alt="Google Scholar" height="60"/>
</a>
<a href="https://lattes.cnpq.br/7025666927284220">
  <img src="https://i.ibb.co/1fMjS38j/lattes.png" alt="Lattes" height="60"/>
</a>

---

> _Feel free to open issues or PRs, or reach out for collaboration or questions!_
