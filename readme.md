# Make-Biz-Plan

This script transforms an Excel export from Viva Goals into a PowerPoint file with one slide per Viva Goals object.

## Description

The script reads data from an Excel workbook, processes the data to create slides in a PowerPoint presentation, and saves the resulting presentation to a specified file. The script supports various command-line arguments to specify the source workbook, template PowerPoint, target PowerPoint, and slide master layouts.

## Installation

1. **Clone the repository** (if applicable):
   ```sh
   git clone <repository_url>
   cd <repository_directory>
   ```

2. **Create a virtual environment** (optional but recommended):
   ```sh
   python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

3. **Install the dependencies**:
   ```sh
   pip install -r requirements.txt
   ```

## Usage

Run the script with the required arguments:

```sh
python Make-Biz-Plan.py --source_workbook VivaGoals.xlsx --template_powerpoint template.pptx --target_bizplan_powerpoint bizplan.pptx
```

### Command-Line Arguments

- `--source_workbook`: Path to the source Excel workbook. Default is 

VivaGoals.xlsx

.
- `--template_powerpoint`: Path to the template PowerPoint file. Default is 

template.pptx

.
- `--target_bizplan_powerpoint`: Path to the target PowerPoint file. Default is 

bizplan.pptx

.
- `--theme_slide_master`: Index of the theme slide master. Default is `0`.
- `--theme_slide_master_layout`: Index of the theme slide master layout. Default is `3`.
- `--okr_slide_master`: Index of the OKR slide master. Default is `2`.
- `--okr_slide_master_layout`: Index of the OKR slide master layout. Default is `11`.

### Example

```sh
python Make-Biz-Plan.py --source_workbook VivaGoals.xlsx --template_powerpoint template.pptx --target_bizplan_powerpoint bizplan.pptx --theme_slide_master 0 --theme_slide_master_layout 3 --okr_slide_master 2 --okr_slide_master_layout 11
```

## Development

### Setting Up the Development Environment

1. **Create a virtual environment**:
   ```sh
   python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

2. **Install the dependencies**:
   ```sh
   pip install -r requirements.txt
   ```

### Running Tests

To run tests, use the following command:

```sh
pytest
```

## Contributing

1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Make your changes.
4. Commit your changes (`git commit -am 'Add new feature'`).
5. Push to the branch (`git push origin feature-branch`).
6. Create a new Pull Request.

## License

This project is licensed under the MIT License. See the LICENSE file for details.
