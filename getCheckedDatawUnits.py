import logging
import clr

# Add the necessary references
clr.AddReference("System")
clr.AddReference(r"C:\Program Files\Audio Precision\APx500 8.0\API\AudioPrecision.API2.dll")
clr.AddReference(r"C:\Program Files\Audio Precision\APx500 8.0\API\AudioPrecision.API.dll")

from AudioPrecision.API import *

class APxDataUnitsRetriever:
    def __init__(self):
        self.APx = APx500_Application()

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
                            #'data': self.process_measurement_data(sequence_result),
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

# Instantiate APxDataUnitsRetriever without arguments
apx_retriever = APxDataUnitsRetriever()
apx_retriever.retrieve_and_print_units()
