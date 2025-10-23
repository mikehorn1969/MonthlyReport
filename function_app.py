import azure.functions as func
import logging
import json
from parse_reports import process_sharepoint_files

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

@app.route(route="http_trigger")
def http_trigger(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Monthly Reports processed a request.')

    filecount = process_sharepoint_files()   
    logging.info(f'Processed {filecount} files.')

    return func.HttpResponse(
        json.dumps({"message": f"Processed {filecount} files."}),
        mimetype="application/json"
    )