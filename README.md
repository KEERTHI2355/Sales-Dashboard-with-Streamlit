# 📊 Sales Insights Dashboard

The **Sales Insights Dashboard** is an end-to-end data analysis and visualization project built using **Python, Pandas, Matplotlib/Seaborn, and Streamlit**, designed to help businesses understand sales performance across products, categories, and regions through a clean dataset, exploratory data analysis (EDA), KPI computation, and an interactive dashboard that highlights best-selling products, category-wise contributions, regional performance, and trends over time, while also providing **Excel export functionality** for easy summary report downloads and showcasing a strong data analytics workflow suitable for **Data Analyst, BI Analyst, and Python-based roles**.


## 🧭 Overview  

This is a web-based dashboard built using **Streamlit**, designed to analyze and visualize sales data.  

It includes:  
- Data loading and cleaning  
- Exploratory data analysis  
- KPI generation  
- Visual charts  
- Interactive dashboard with filters  
- Export-to-Excel functionality  


## 🎯 Objectives  

The main objective of the **Sales Insights Dashboard** is to provide a structured, analytical workflow that converts raw sales data into meaningful visual insights.  

Specific objectives include:  

**1. Data Cleaning and Structuring**  
Ensure the raw sales dataset is cleaned, well-structured, and ready for analysis.  

**2. Exploratory Data Analysis (EDA)**  
Generate charts showing product sales trends, category distribution, and regional performance.  

**3. KPI Computation**  
Calculate key performance indicators such as total revenue, order volume, average order value, and top-performing regions/products.  

**4. Interactive Dashboard**  
Create an easy-to-use Streamlit application allowing users to explore sales insights with filters.  

**5. Excel Export Feature**  
Enable users to download summarized reports for business usage or presentations.  

**6. Reusability & Scalability**  
Structure the project so it can be easily extended to new datasets or integrated with APIs.  


## 🗂️ Files and Folders  

`
Generate_Sales_Dataset.py
` - Generates a synthetic sales dataset for analysis, including fields such as order date, product, category, region, quantity, price, and sales.  

`
sales_dataset.csv
` - The raw dataset used for EDA and dashboard creation.  

`
sales_cleaned.csv
` - A cleaned version of the dataset generated after data preprocessing.  

`
notebooks/
` - This folder contains all the Jupyter notebooks used for analysis.  

`
01_eda.ipynb
` - Initial exploration: shape, datatypes, missing values.  

`
02_eda.ipynb
` - Feature engineering (e.g., extracting order month), aggregations, and summary statistics.  

`
03_eda(plots).ipynb
` - Chart creation: product sales, category share, and regional performance.  

`
app/app.py
` -The main Streamlit application file.  
<pre> - Contains:  
		- Sidebar filters  
		- KPI cards  
		- Chart rendering  
		- Excel export button  
</pre>


`
charts/
` - Stores chart images generated during EDA.  
<pre> - Contains
		- product_sales.png  
		- category_sales.png  
		- region_sales.png  </pre>

`
KPI.py
` - Python file containing KPI calculations (total revenue, top product, etc.).  


## 📁 File Structure  

```
Sales Insights Dashboard/
  data/
     Generate_Sales_Dataset.py
     ales_cleaned.csv
     sales_dataset.csv  
  notebooks/
     01_eda.ipynb
     02_eda.ipynb
     03_eda(plots).ipynb    
  app/
     app.py   
  charts/
     product_sales.png
     category_sales.png
     region_sales.png
  KPI.py
  README.md
```


## 🚀 How to Run  

**1. Install dependencies**  

```bash
pip install -r requirements.txt
```

**2. Run the Streamlit application**  

```bash
streamlit run app/app.py
```

**3. The dashboard will open in your browser automatically.**  


## 🧪 How to Use  

- Choose filters (month, region, category) from the sidebar.  
- View KPIs at the top of the dashboard.  
- Scroll to see charts such as product sales and category-wise share.  
- Click **Download Excel Report** to export summarized KPIs and tables.  


## ⚙️ Prerequisites  

- Python 3.8 or higher  
- Streamlit 1.25+  
- Pandas, NumPy, Matplotlib, Seaborn  
- A modern web browser (Chrome, Edge, Firefox)  


## 🛠 Installation  

1. Install Python from the official website.  

2. Install Streamlit via:  
   ```bash
   pip install streamlit
   ```

3. Clone the repository:  
   ```bash
   git clone https://github.com/KEERTHI2355/Sales-Insights-Dashboard
   ```

4. Navigate into the project:  
   ```bash
   cd Sales-Insights-Dashboard
   ```

5. Install dependencies:  
   ```bash
   pip install -r requirements.txt
   ```


## 🤝 How to Contribute  

Contributions, feature requests, and issues are welcome!  

1. Fork the repository  

2. Create your feature branch  
   ```bash
   git checkout -b feature/new-feature
   ```

3. Commit your changes  
   ```bash
   git commit -m "Add new feature"
   ```

4. Push the branch  
   ```bash
   git push origin feature/new-feature
   ```

5. Open a pull request  


## 📄 License  

This project is licensed under the **MIT License** — see the `LICENSE` file for details.  


## 📝 Note  

- Replace the synthetic dataset with real business data if available.  
- Review and adjust chart formatting inside `app.py` based on your visual preferences.  
- Ensure `xlsxwriter` is installed for Excel export functionality.


## 👤 Author

**K Keerthi**  
Data Science Engineering Student  
Aspiring Python Developer / Data Analyst  

📌 *This project was built end-to-end with an industry-first mindset.*

---

<p align="center">
  🛠 Built by <a href="https://github.com/KEERTHI2355">@Keerthi2355</a> 
  <br>
  <a href="#-the-vault">Back to Top</a>
</p>

