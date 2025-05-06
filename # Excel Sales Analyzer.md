# Excel Sales Analyzer

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)

A Python tool for generating sample sales data, analyzing it, and creating professional Excel reports with charts and visualizations.

![Sample Report Preview](https://via.placeholder.com/800x400?text=Sample+Excel+Report)

## Features

- 📊 Generates realistic sales data with configurable parameters
- 📑 Creates formatted Excel files with proper styling
- 📈 Performs comprehensive sales analysis:
  - Product performance metrics
  - Regional sales breakdown
  - Sales channel effectiveness
  - Time-based trend analysis
- 📝 Builds multi-sheet summary reports with charts and visualizations
- 🔄 Easy to integrate into existing data pipelines

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/excel-sales-analyzer.git
cd excel-sales-analyzer
```

2. Create a virtual environment (optional but recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

Run the program with default settings:

```bash
python sales_analyzer.py
```

This will:
1. Generate 150 sample sales records
2. Save them to `output/sales_data.xlsx`
3. Analyze the data
4. Create a summary report at `output/sales_summary.xlsx`

### Command Line Options

```
usage: sales_analyzer.py [-h] [--records RECORDS] [--days DAYS] [--output-dir OUTPUT_DIR] 
                        [--data-file DATA_FILE] [--report-file REPORT_FILE] [--verbose]

Excel Sales Data Generator and Analyzer

optional arguments:
  -h, --help            show this help message and exit
  --records RECORDS     Number of sample records to generate
  --days DAYS           Number of days in the past to generate data for
  --output-dir OUTPUT_DIR
                        Directory to save output files
  --data-file DATA_FILE
                        Filename for raw data
  --report-file REPORT_FILE
                        Filename for summary report
  --verbose             Enable verbose logging
```

### Examples

Generate 500 sales records over a 180-day period:
```bash
python sales_analyzer.py --records 500 --days 180
```

Specify custom output directory and filenames:
```bash
python sales_analyzer.py --output-dir "my_reports" --data-file "my_sales.xlsx" --report-file "my_analysis.xlsx"
```

Enable verbose logging:
```bash
python sales_analyzer.py --verbose
```

## Using as a Module

You can also use the components in your own Python code:

```python
from sales_analyzer import SalesDataGenerator, ExcelWriter, SalesAnalyzer, ReportGenerator

# Generate sample data
generator = SalesDataGenerator()
sales_df = generator.generate_data(num_records=200)

# Write to Excel
writer = ExcelWriter()
sales_file = writer.write_sales_data(sales_df, filename="my_sales.xlsx")

# Analyze data
analyzer = SalesAnalyzer(sales_df)
results = analyzer.analyze()

# Create report
report_gen = ReportGenerator()
report_file = report_gen.create_report(results, filename="my_report.xlsx")

print(f"Report generated at: {report_file}")
```

## Structure

```
excel-sales-analyzer/
├── sales_analyzer.py      # Main script
├── requirements.txt       # Dependencies
├── README.md              # This file
├── LICENSE                # MIT License
└── output/                # Default output directory
    ├── sales_data.xlsx    # Generated sample data
    └── sales_summary.xlsx # Analysis report
```

## Dependencies

- pandas: Data manipulation and analysis
- numpy: Numerical operations
- matplotlib: Visualization support
- xlsxwriter: Excel file creation with advanced formatting
- openpyxl: Excel file reading

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Thanks to all the open-source libraries that made this project possible
- Inspired by real-world business analytics needs