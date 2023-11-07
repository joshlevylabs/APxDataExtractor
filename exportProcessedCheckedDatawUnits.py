import logging
import clr
import openpyxl
from openpyxl import Workbook
import argparse
import sys

# Set up the argument parser
parser = argparse.ArgumentParser(description='Export data to Excel with additional unit name.')
parser.add_argument('-name', '--name', help='Name of the unit under test to append to the Excel file.', default='')

# Parse command-line arguments
args = parser.parse_args()

# Now args.name contains the name provided by the user or an empty string if not provided
dut_name = args.name

# Add the necessary references
clr.AddReference("System")
clr.AddReference(r"C:\Program Files\Audio Precision\APx500 8.0\API\AudioPrecision.API2.dll")
clr.AddReference(r"C:\Program Files\Audio Precision\APx500 8.0\API\AudioPrecision.API.dll")

from AudioPrecision.API import *

class APxDataUnitsRetriever:
    def __init__(self, dut_name):
        self.APx = APx500_Application()
        self.dut_name = dut_name

    def get_xy_units(self, signal_path_index, measurement_index, result_index):
        measurement = self.APx.Sequence.GetMeasurement(signal_path_index, measurement_index)
        sequence_result = measurement.SequenceResults[result_index]

        x_unit = 'N/A'
        y_unit = 'N/A'

        # Get XUnit
        try:
            x_unit = sequence_result.XUnit
        except Exception as e:
            logging.error(f"Error accessing XUnit for Signal Path {signal_path_index}, Measurement {measurement_index}, Result {result_index}: {e}")

        # Try to get YUnit, handle exception if not available
        try:
            y_unit = sequence_result.YUnit
        except Exception as e:
            logging.info(f"YUnit not available for Signal Path {signal_path_index}, Measurement {measurement_index}, Result {result_index}: {e}")

        return {"x_unit": x_unit, "y_unit": y_unit}

    def retrieve_and_print_units(self):
        checked_signal_paths = self.GetCheckedData()

        for sp_data in checked_signal_paths:
            sp_name = sp_data['name']  # Get the signal path name
            sp_idx = sp_data['index']

            for m_data in sp_data['measurements']:
                m_name = m_data['name']  # Get the measurement name
                m_idx = m_data['index']

                for result_data in m_data['results']:
                    r_name = result_data['name']  # Get the result name
                    result_idx = result_data['index']
                    passed = "PASS" if result_data['passed'] else "FAIL"  # Determine Pass/Fail
                    units = result_data['units']
                    data = result_data['data']  # Get the processed measurement data

                    # Prepare the data string for printing, handling the potential complexity of data content
                    data_string = ", ".join(f"{key}: {value}" for key, value in data.items())

                    # Print out the desired information
                    print(f"Signal Path [{sp_idx}]: {sp_name}, Measurement [{m_idx}]: {m_name}, "
                        f"Result [{result_idx}]: {r_name} - Status: {passed}, "
                        f"XUnit: {units['x_unit']}, YUnit: {units['y_unit']}")

    def GetCheckedData(self, sender=None, args=None):
        if not self.APx:
            logging.error("APx is None. Launch AP Software first.")
            return

        checked_signal_paths = []
        try:
            for sp_idx, sp in enumerate(self.APx.Sequence):
                signal_path = ISignalPath(sp)
                if not signal_path.Checked:
                    continue
                current_sp = {"name": signal_path.Name, "measurements": [], "index": sp_idx}  # Include the index here
                logging.info(f"Checked Signal Path: {signal_path.Name}")

                for m_idx, m in enumerate(signal_path):
                    measurement = ISequenceMeasurement(m)
                    if not measurement.Checked:
                        continue
                    current_measurement = {"name": measurement.Name, "index": m_idx, "results": []}  # Include the index here
                    logging.info(f"\tChecked Measurement: {measurement.Name}")

                    for result_idx, result in enumerate(measurement.SequenceResults):
                        sequence_result = ISequenceResult(result)

                        # Retrieve units
                        units = self.get_xy_units(sp_idx, m_idx, result_idx)

                        failed = self.is_result_failed(signal_path.Name, measurement.Name, sequence_result.Name)
                        error = self.result_error(signal_path.Name, measurement.Name, sequence_result.Name)
                        status = "ERROR" if error else ("FAIL" if failed else "PASS")
                        logging.info(f"{signal_path.Name} | {measurement.Name} | {sequence_result.Name} -> {status}")

                        current_result = {
                            'name': sequence_result.Name,
                            'index': result_idx,  # Store index
                            'result_object': sequence_result,
                            'data': self.process_measurement_data(sequence_result),
                            'passed': not (failed or error),
                            'units': units
                        }
                        current_measurement["results"].append(current_result)

                    current_sp["measurements"].append(current_measurement)

                checked_signal_paths.append(current_sp)
                logging.info(f"Added a new signal path to checked_signal_paths. Total items now: {len(checked_signal_paths)}")

        except IndexError as e:
            logging.error(f"Index out of range error: {e}")
            # Include additional debug information if necessary
            logging.debug("Debug Info: APx.Sequence Length: {}".format(len(self.APx.Sequence)))
        except Exception as e:
            logging.error(f"An unexpected error occurred: {e}")
            logging.exception("Exception Stack Trace")

        return checked_signal_paths
    
    def is_result_failed(self, signal_path_name, measurement_name, result_name):
        try:
            for idx in range(self.APx.Sequence.FailedResults.Count):
                failed_result = self.APx.Sequence.FailedResults[idx]
                if (failed_result.Name == result_name and
                    failed_result.MeasurementName == measurement_name and
                    failed_result.SignalPathName == signal_path_name):
                    return not failed_result.PassedResult
        except Exception as e:
            logging.error(f"Error retrieving failed status for result {result_name}: {str(e)}")
        return False
    
    def result_error(self, signal_path_name, measurement_name, result_name):
        try:
            for idx in range(self.APx.Sequence.Results.Count):
                result = self.APx.Sequence.Results[idx]
                if (result.Name == result_name and
                    result.MeasurementName == measurement_name and
                    result.SignalPathName == signal_path_name):
                    return result.HasErrorMessage
        except Exception as e:
            logging.error(f"Error retrieving error status for result {result_name}: {str(e)}")
        return False
    
    def process_measurement_data(self, sequence_result):
        data = {}
        
        # Check if the result object has the 'PassedLimitChecks' attribute
        if hasattr(sequence_result, 'PassedLimitChecks'):
            passed_limit_checks = sequence_result.PassedLimitChecks
            logging.info(f"\t\tPassedLimitChecks for {sequence_result.Name}: {passed_limit_checks}")
        else:
            passed_limit_checks = False
            #logging.warning(f"\t\t{sequence_result.Name} does not have PassedLimitChecks attribute. Setting to False.")

        # Check and Get XYValues
        for vertical_axis in [VerticalAxis.Left, VerticalAxis.Right]:
            channel_count = sequence_result.GetXYChannelCount(vertical_axis)
            if channel_count > 0:
                xValues, yValues = [], []
                for ch in range(channel_count):
                    graphPoints = sequence_result.GetXYValues(ch, vertical_axis, SourceDataType.Measured, 1)
                    if graphPoints:
                        xVals, yVals = [point.X for point in graphPoints], [point.Y for point in graphPoints]
                        xValues.append(xVals)
                        yValues.append(yVals)
                data['xValues'] = xValues
                data['yValues'] = yValues
        
        # Check for meter values
        if sequence_result.HasMeterValues:
            meterValues = sequence_result.GetMeterValues()
            data['meterValues'] = meterValues
            logging.info(f"\t\tFound Meter Values for Result: {sequence_result.Name}")
            # Log each meter value directly.
            for idx, value in enumerate(meterValues):
                logging.info(f"\t\t\tMeter Value {idx}: {value}")

        # Check and Get RawTextResults
        if sequence_result.HasRawTextResults:
            rawTextResults = sequence_result.GetRawTextResults()
            data['rawTextResults'] = rawTextResults
            logging.info(f"\t\tFound Raw Text Results for Result: {sequence_result.Name}")

        # Check and Get XValues and YValues
        for ch in range(sequence_result.GetXYChannelCount(VerticalAxis.Left)):
            for vertical_axis in [VerticalAxis.Left, VerticalAxis.Right]:
                for data_type in [SourceDataType.Measured]:  # Only considering SourceDataType.Measured as per the older logic
                    if hasattr(sequence_result, 'HasXValues') and sequence_result.HasXValues(ch, vertical_axis, data_type):
                        xValues = sequence_result.GetXValues(ch, vertical_axis, data_type, 1)
                        data[f'xValues_{vertical_axis}_{data_type}'] = xValues
                    if hasattr(sequence_result, 'HasYValues') and sequence_result.HasYValues(ch, vertical_axis, data_type):
                        yValues = sequence_result.GetYValues(ch, vertical_axis, data_type, 1)
                        data[f'yValues_{vertical_axis}_{data_type}'] = yValues

        return data
    
    ####### Export to excel section #######

    def export_to_excel(self, checked_signal_paths):
        workbook = Workbook()
        active_sheet = workbook.active

        # Log the checked_signal_paths structure for debugging
        logging.info(f"Checked Signal Paths: {checked_signal_paths}")

        for sp_data in checked_signal_paths:
            logging.info(f"Processing signal path: {sp_data['name']}")
            for m_data in sp_data['measurements']:
                logging.info(f"Processing measurement: {m_data['name']}")
                for result_data in m_data['results']:
                    try:
                        worksheet = workbook.create_sheet(title=result_data['name'])
                    except Exception as e:
                        logging.error(f"Could not create a sheet with name '{result_data['name']}': {e}")
                        continue

                    headers = ['Signal Path', 'Measurement', 'Result Index', 'Status', 'XUnit', 'YUnit']
                    worksheet.append(headers)
                    logging.info(f"Headers appended for sheet '{result_data['name']}'")

                    # Additional error handling for unexpected structures
                    try:
                        if 'xyvalues' in result_data:
                            for i, (x, y) in enumerate(zip(result_data['xyvalues']['x'], result_data['xyvalues']['y'])):
                                data_row = [sp_data['name'], m_data['name'], f"{result_data['name']} (Pair {i+1})",
                                            "PASS" if result_data['passed'] else "FAIL",
                                            result_data['units']['x_unit'], result_data['units']['y_unit'], x, y]
                                worksheet.append(data_row)
                            logging.info(f"Appended {len(result_data['xyvalues']['x'])} rows of xyvalues to '{result_data['name']}'")

                        elif 'xvalues' in result_data:
                            for i, x in enumerate(result_data['xvalues']):
                                data_row = [sp_data['name'], m_data['name'], f"{result_data['name']} (X {i+1})",
                                            "PASS" if result_data['passed'] else "FAIL",
                                            result_data['units']['x_unit'], '', x]
                                worksheet.append(data_row)
                            logging.info(f"Appended {len(result_data['xvalues'])} rows of xvalues to '{result_data['name']}'")

                        elif 'metervalues' in result_data:
                            for i, meter_value in enumerate(result_data['metervalues']):
                                data_row = [sp_data['name'], m_data['name'], f"{result_data['name']} (Meter {i+1})",
                                            "PASS" if result_data['passed'] else "FAIL", '', '', meter_value]
                                worksheet.append(data_row)
                            logging.info(f"Appended {len(result_data['metervalues'])} rows of metervalues to '{result_data['name']}'")

                        else:
                            logging.error(f"Unknown data structure for result {result_data['name']}: {result_data}")
                            logging.debug(f"Result data structure: {result_data}")  # Log the unknown structure

                    except Exception as e:
                        logging.error(f"An error occurred while processing result {result_data['name']}: {e}")
                        logging.debug(traceback.format_exc())  # Log the complete traceback

        if not active_sheet['A1'].value:
            workbook.remove(active_sheet)

        # Make sure self.dut_name is not None or empty
        excel_filename = f"{self.dut_name or 'ExportedData'}.xlsx"
        workbook.save(excel_filename)
        logging.info(f"Workbook saved as '{excel_filename}'")

# Initialize logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Instantiate APxDataUnitsRetriever with 'Unit under test' name argument
apx_retriever = APxDataUnitsRetriever(dut_name='UnitUnderTestName')
try:
    checked_signal_paths = apx_retriever.GetCheckedData()
    apx_retriever.export_to_excel(checked_signal_paths)
except Exception as e:
    logging.error(f"An error occurred: {e}")