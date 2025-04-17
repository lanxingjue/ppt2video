import logging
import shutil
from pathlib import Path
import time

# --- Configure Logging ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s', # Include filename
    datefmt='%Y-%m-%d %H:%M:%S'
)

# --- Import functions from our modules ---
try:
    # ppt_processor now implicitly handles ppt_exporter_win
    from ppt_processor import process_presentation
    logging.info("Successfully imported 'process_presentation' from ppt_processor.")
except ImportError:
    logging.error("Failed to import 'process_presentation' from ppt_processor.py. Ensure the file exists and has no errors.", exc_info=True)
    exit()
except Exception as e:
    logging.error(f"An unexpected error occurred during import from ppt_processor: {e}", exc_info=True)
    exit()

try:
    from video_synthesizer import create_video_from_data
    logging.info("Successfully imported 'create_video_from_data' from video_synthesizer.")
except ImportError:
     logging.error("Failed to import 'create_video_from_data' from video_synthesizer.py. Ensure the file exists and has no errors.", exc_info=True)
     exit()
except Exception as e:
    logging.error(f"An unexpected error occurred during import from video_synthesizer: {e}", exc_info=True)
    exit()


# --- Configuration ---
# !!! PLEASE MODIFY THESE PATHS !!!
INPUT_PPTX_FILE = Path("智能短信分类平台方案.pptx")  # <--- Set your input PPTX file path here
# Base directory where output video and temporary folders will be created
BASE_OUTPUT_DIR = Path("./full_process_output")
# Set to True to delete the temporary folder after successful completion
CLEANUP_TEMP_DIR = True
# !!! --- End of Configuration --- !!!


def run_full_process():
    """
    Executes the entire PPT to Video conversion process.
    """
    start_time = time.time()
    logging.info("="*20 + " Starting Full PPT to Video Process " + "="*20)

    # 1. Validate Input
    if not INPUT_PPTX_FILE.is_file():
        logging.error(f"Input PPTX file not found: {INPUT_PPTX_FILE}")
        return

    # 2. Ensure Base Output Directory Exists
    try:
        BASE_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        logging.info(f"Base output directory ensured: {BASE_OUTPUT_DIR.resolve()}")
    except OSError as e:
        logging.error(f"Failed to create or access base output directory '{BASE_OUTPUT_DIR}': {e}")
        return

    # 3. Define Final Output Video Path
    final_video_filename = INPUT_PPTX_FILE.stem + "_final_video.mp4"
    final_video_path = BASE_OUTPUT_DIR / final_video_filename
    # Ensure target video file doesn't exist initially if we use shutil.move later
    if final_video_path.exists():
        logging.warning(f"Output video file already exists: {final_video_path}. It will be overwritten.")
        try:
            final_video_path.unlink()
        except OSError as e:
             logging.error(f"Failed to delete existing output file '{final_video_path}': {e}. Please remove it manually.")
             return


    # --- === Step 1 & 2: Process Presentation (Export Images, Extract Notes, Generate Audio) === ---
    logging.info("--- Running Step: Processing Presentation (Export, Notes, Audio)... ---")
    processed_result_tuple = None # Initialize
    temp_run_dir = None         # Initialize

    try:
        processed_result_tuple = process_presentation(INPUT_PPTX_FILE, BASE_OUTPUT_DIR)
    except Exception as e:
        # Catch potential errors during the call itself (though the function should handle internal errors)
        logging.error(f"An unexpected error occurred while calling 'process_presentation': {e}", exc_info=True)
        # Attempt to clean up potential partial temp dir if path was assigned? Unlikely here.
        return # Stop processing

    # Check the result from process_presentation
    if processed_result_tuple and isinstance(processed_result_tuple, tuple) and len(processed_result_tuple) == 2:
        processed_data, temp_run_dir = processed_result_tuple
        if processed_data is not None and temp_run_dir is not None:
            logging.info(f"Presentation processing completed. Data for {len(processed_data)} slides processed.")
            logging.info(f"Temporary files located in: {temp_run_dir.resolve()}")
        else:
            logging.error("Presentation processing function returned invalid data or temp directory path. Aborting.")
            # Should we attempt cleanup of temp_run_dir if it's not None but data is?
            if temp_run_dir and temp_run_dir.exists() and CLEANUP_TEMP_DIR:
                logging.warning(f"Attempting to clean up potentially incomplete temp directory: {temp_run_dir}")
                try: shutil.rmtree(temp_run_dir)
                except Exception as e: logging.error(f"Cleanup failed: {e}")
            return # Stop processing
    else:
        logging.error("Presentation processing failed or returned unexpected result. Check logs from 'ppt_processor'. Aborting.")
        # No reliable temp_run_dir path here.
        return # Stop processing


    # --- === Step 3: Synthesize Video (Combine Segments, Add Subtitles) === ---
    logging.info("--- Running Step: Synthesizing Video... ---")
    synthesis_success = False
    try:
        synthesis_success = create_video_from_data(
            processed_data,
            temp_run_dir, # Pass the temp directory path received from step 1/2
            final_video_path
        )
    except Exception as e:
        # Catch potential errors during the call itself
        logging.error(f"An unexpected error occurred while calling 'create_video_from_data': {e}", exc_info=True)
        synthesis_success = False # Ensure it's marked as failed

    # --- === Step 4: Final Output and Cleanup === ---
    if synthesis_success and final_video_path.exists():
        end_time = time.time()
        logging.info("="*20 + " Process Completed Successfully! " + "="*20)
        logging.info(f"Final video saved to: {final_video_path.resolve()}")
        logging.info(f"Total processing time: {end_time - start_time:.2f} seconds")

        # Optional Cleanup
        if CLEANUP_TEMP_DIR:
            logging.info(f"Attempting to clean up temporary directory: {temp_run_dir}")
            try:
                shutil.rmtree(temp_run_dir)
                logging.info("Temporary directory successfully cleaned up.")
            except Exception as e:
                logging.warning(f"Failed to clean up temporary directory '{temp_run_dir}': {e}")
        else:
            logging.info(f"Temporary files retained in: {temp_run_dir}")

    else:
        end_time = time.time()
        logging.error("="*20 + " Process Failed! " + "="*20)
        logging.error("Video synthesis failed or the output file was not created.")
        logging.error(f"Total processing time: {end_time - start_time:.2f} seconds")
        if temp_run_dir and temp_run_dir.exists():
            logging.info(f"Temporary files related to this run are kept for inspection in: {temp_run_dir.resolve()}")
        else:
             logging.info("No temporary directory path available or directory does not exist.")


# --- Run the main process ---
if __name__ == "__main__":
    run_full_process()