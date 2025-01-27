from flask import Flask, request, jsonify
from process_excel import process_excel
import os

app = Flask(__name__)

@app.route('/process', methods=['POST'])
def process_files():
    try:
        # Get uploaded files
        excel1 = request.files['excel1']
        excel2 = request.files['excel2']

        # Save uploaded files temporarily
        os.makedirs('temp', exist_ok=True)
        excel1_path = f"./temp/{excel1.filename}"
        excel2_path = f"./temp/{excel2.filename}"
        excel1.save(excel1_path)
        excel2.save(excel2_path)

        # Process files
        output_path, missing_data_path = process_excel(excel1_path, excel2_path)

        # Return the result
        return jsonify({
            "message": "Processing completed.",
            "output_path": output_path,
            "missing_data_path": missing_data_path or "No missing data."
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000)
