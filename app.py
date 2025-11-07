# app.py
from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd
import io
import contextlib
import traceback

app = Flask(__name__)
# Enable Cross-Origin Resource Sharing (CORS) to allow your frontend to call the backend
CORS(app)

@app.route('/execute', methods=['POST'])
def execute_code():
    try:
        data = request.get_json()
        code = data.get('code')
        sheets_data = data.get('sheets')
        active_sheet_name = data.get('activeSheetName')

        if not code or not sheets_data:
            return jsonify({"error": "Missing 'code' or 'sheets' in request."}), 400

        # Convert the incoming sheet JSON data into a dictionary of pandas DataFrames
        # This is what the AI's generated Python code will expect
        dfs = {}
        for sheet in sheets_data:
            df = pd.DataFrame(sheet['cells'])
            # Use the first row as headers if they exist, otherwise use standard A, B, C...
            if not df.empty:
                df.columns = [chr(65 + i) for i in range(len(df.columns))]
            dfs[sheet['name']] = df

        # Prepare a dictionary to be used as the local scope for exec()
        # It includes the 'dfs' dictionary and pandas library
        local_scope = {'dfs': dfs, 'pd': pd}

        # Use contextlib to capture stdout and stderr from the executed code
        stdout_capture = io.StringIO()
        stderr_capture = io.StringIO()

        with contextlib.redirect_stdout(stdout_capture), contextlib.redirect_stderr(stderr_capture):
            exec(code, {'__builtins__': __builtins__}, local_scope)

        # After execution, retrieve the (potentially modified) DataFrames
        final_dfs = local_scope['dfs']

        # Convert the DataFrames back into the JSON structure the frontend expects
        updated_sheets_data = []
        for original_sheet in sheets_data:
            sheet_name = original_sheet['name']
            if sheet_name in final_dfs:
                df = final_dfs[sheet_name]
                # Convert DataFrame back to a simple list of lists for the 'cells'
                # and fill NaN with empty strings
                new_cells = df.fillna('').values.tolist()
                
                # Make sure the sheet dimensions match the original or are larger
                original_row_count = len(original_sheet['cells'])
                original_col_count = len(original_sheet['cells'][0]) if original_row_count > 0 else 0
                
                new_row_count = len(new_cells)
                new_col_count = len(new_cells[0]) if new_row_count > 0 else 0
                
                final_row_count = max(original_row_count, new_row_count)
                final_col_count = max(original_col_count, new_col_count)

                # Create a correctly sized grid filled with empty strings
                final_cells = [['' for _ in range(final_col_count)] for _ in range(final_row_count)]

                # Populate it with the new data
                for r_idx, row in enumerate(new_cells):
                    for c_idx, cell in enumerate(row):
                        # Ensure cells are strings for JSON serialization
                        final_cells[r_idx][c_idx] = str(cell)
                
                original_sheet['cells'] = final_cells
                # You could also update columnWidths/rowHeights here if needed
                original_sheet['columnWidths'] = [original_sheet['columnWidths'][0]] * final_col_count
                original_sheet['rowHeights'] = [original_sheet['rowHeights'][0]] * final_row_count
                original_sheet['formats'] = [[{} for _ in range(final_col_count)] for _ in range(final_row_count)]


            updated_sheets_data.append(original_sheet)

        return jsonify({"sheets": updated_sheets_data})

    except Exception as e:
        # If any error occurs during execution, capture it and send it back to the frontend
        tb_str = traceback.format_exc()
        return jsonify({"error": f"Python execution error:\n{tb_str}"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5001)
