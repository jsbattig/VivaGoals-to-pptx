import unittest
from unittest.mock import patch, MagicMock
from Make_Biz_Plan import *

class MockWorkbook:
    def __init__(self, test_data):
        self.active = MockWorksheet(test_data)
        self.active.iter_rows = self.active.iter_rows  # Ensure method is accessible

class MockWorksheet:
    def __init__(self, test_data):
        self.test_data = test_data
        self._current_row = 0

    def iter_rows(self, min_row=1, max_row=None, values_only=False):
        for row in self.test_data[min_row-1:]:
            yield row if values_only else [MagicMock(value=cell) for cell in row]

class MockPresentation:
    def __init__(self):
        self.slides = MockSlideCollection()
        self.slide_masters = [
            MagicMock(slide_layouts=[MagicMock() for _ in range(12)])
            for _ in range(3)
        ]
        self.saved_file = None

    def save(self, filename):
        self.saved_file = filename

class MockSlideCollection:
    def __init__(self):
        self.slides = []

    def add_slide(self, layout):
        slide = MagicMock()
        # Create a new shape collection for each slide
        shape_collection = MockShapeCollection()
        shape_collection.title = MagicMock()  # Add title attribute
        slide.shapes = shape_collection
        slide.layout = layout
        self.slides.append(slide)
        return slide

class MockShapeCollection:
    def __init__(self):
        self.shapes = []
        self._spTree = MagicMock()
        self.title = MagicMock()
        self.text_frame = MagicMock()
        # Add an initial shape with text_frame for negative indexing
        initial_shape = MagicMock()
        initial_shape.text_frame = MagicMock()
        self.shapes.append(initial_shape)

    def __getitem__(self, key):
        if isinstance(key, int):
            if key < 0:
                key = len(self.shapes) + key
            if 0 <= key < len(self.shapes):
                return self.shapes[key]
            raise IndexError("Shape index out of range")
        return self.shapes[key]

    def __setitem__(self, key, value):
        if isinstance(key, int):
            if key < 0:
                key = len(self.shapes) + key
            while len(self.shapes) <= key:
                self.shapes.append(MagicMock())
            self.shapes[key] = value

    def __len__(self):
        return len(self.shapes)

    def add_textbox(self, *args):
        textbox = MagicMock()
        textbox.text_frame = MagicMock()
        self.shapes.append(textbox)
        return textbox

    def add_picture(self, *args):
        picture = MagicMock()
        picture.click_action = MagicMock()
        picture.click_action.hyperlink = MagicMock()
        self.shapes.append(picture)
        return picture

    def add_shape(self, *args):
        shape = MagicMock()
        shape.fill = MagicMock()
        shape.line = MagicMock()
        shape._element = MagicMock()
        self.shapes.append(shape)
        return shape

    def add_connector(self, *args):
        connector = MagicMock()
        self.shapes.append(connector)
        return connector

class TestEndToEnd(unittest.TestCase):
    def setUp(self):
        # Test data setup
        self.headers = ['Id', 'Title', 'Tag', 'Owner', 'Period', 'Start Date', 'End Date',
                       'Description', 'Aligned To (weight, Objective ID)', 'Metric Name',
                       'Target', 'Object Type', 'Status']
        
        self.test_data = [
            self.headers,
            ['"http://example.com/1" "1"', 'Theme 1', 'Theme', 'John', 'Q1', '2024-01-01', '2024-03-31',
             'Theme Description', '', 'Metric1', '100%', 'Objective', 'On Track'],
            ['"http://example.com/2" "2"', 'Objective 1', '', 'Jane', 'Q1', '2024-01-01', '2024-03-31',
             'Objective Description', '(weight: 100%, Id: 1)', 'Metric2', '50%', 'Objective', 'At Risk'],
            ['"http://example.com/3" "3"', 'Action 1', '', 'Bob', 'Q1', '2024-01-01', '2024-03-31',
             'Action Description', '(weight: 100%, Id: 2)', 'Metric3', '75%', 'Action', 'On Track']
        ]

        # Initialize mocks correctly
        self.mock_wb = MockWorkbook(self.test_data)
        self.mock_prs = MockPresentation()
        
        # Store original functions and data
        import Make_Biz_Plan
        self.original_goals_dict = Make_Biz_Plan.goals_dict.copy()
        self.original_get_workbook = Make_Biz_Plan.get_workbook
        Make_Biz_Plan.goals_dict = {}

    def tearDown(self):
        # Restore original goals_dict
        import Make_Biz_Plan
        Make_Biz_Plan.goals_dict = self.original_goals_dict
        Make_Biz_Plan.get_workbook = self.original_get_workbook

    @patch('Make_Biz_Plan.Presentation')
    @patch('Make_Biz_Plan.get_workbook')
    def test_end_to_end_flow(self, mock_get_workbook, mock_presentation):
        # Setup mocks in correct order
        mock_get_workbook.return_value = self.mock_wb
        mock_presentation.return_value = self.mock_prs

        # Run main function
        main(source_workbook='test.xlsx',
             template_powerpoint='template.pptx',
             target_bizplan_powerpoint='test_output.pptx')

        # Verify mocks were called
        mock_get_workbook.assert_called_once_with('test.xlsx')
        mock_presentation.assert_called_once_with('template.pptx')

        # Verify presentation was created and saved
        self.assertEqual(self.mock_prs.saved_file, 'test_output.pptx')
        
        # Get all slide titles
        slides = self.mock_prs.slides.slides
        actual_titles = [slide.shapes.title.text for slide in slides]
        expected_titles = ['Theme 1', 'Objective 1', 'Action 1']
        
        self.assertEqual(len(slides), len(expected_titles))
        self.assertEqual(actual_titles, expected_titles)

    @patch('Make_Biz_Plan.Presentation')
    @patch('Make_Biz_Plan.get_workbook')
    def test_correct_slide_ordering(self, mock_get_workbook, mock_presentation):
        """Test that slides are created in the correct order based on dependencies"""
        # Add an outcome to test ordering
        test_data_with_outcome = self.test_data + [
            ['"http://example.com/4" "4"', 'Outcome 1', '', 'Alice', 'Q1', '2024-01-01', '2024-03-31',
             'Outcome Description', '(weight: 100%, Id: 2)', 'Metric4', '25%', 'Outcome', 'On Track']
        ]
        mock_wb = MockWorkbook(test_data_with_outcome)
        mock_prs = MockPresentation()
        
        mock_get_workbook.return_value = mock_wb
        mock_presentation.return_value = mock_prs

        # Run main function
        main(source_workbook='test.xlsx',
             template_powerpoint='template.pptx',
             target_bizplan_powerpoint='test_output.pptx')

        # Verify slides are in correct order
        slides = mock_prs.slides.slides
        expected_titles = ['Theme 1', 'Objective 1', 'Outcome 1', 'Action 1']
        actual_titles = [slide.shapes.title.text for slide in slides]
        
        self.assertEqual(len(slides), len(expected_titles))
        self.assertEqual(actual_titles, expected_titles)

    @patch('Make_Biz_Plan.Presentation')
    @patch('Make_Biz_Plan.get_workbook')
    def test_error_handling(self, mock_get_workbook, mock_presentation):
        """Test error handling for invalid input data"""
        # Test data with invalid object type
        invalid_data = self.test_data + [
            ['"http://example.com/5" "5"', 'Invalid Goal', '', 'Eve', 'Q1', 
             '2024-01-01', '2024-03-31', 'Invalid Description', '', 
             'Metric5', '0%', 'InvalidType', 'On Track']
        ]
        
        mock_wb = MockWorkbook(invalid_data)
        mock_prs = MockPresentation()
        
        # Setup mocks in correct order
        mock_get_workbook.return_value = mock_wb
        mock_presentation.return_value = mock_prs

        # Run main function and verify error is raised with correct message
        with self.assertRaises(ValueError) as context:
            main(source_workbook='test.xlsx',
                 template_powerpoint='template.pptx',
                 target_bizplan_powerpoint='test_output.pptx')

        self.assertIn("Invalid object type", str(context.exception))

if __name__ == '__main__':
    unittest.main()
