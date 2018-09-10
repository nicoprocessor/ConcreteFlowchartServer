import pathlib

from flask import Flask, request, jsonify, send_from_directory, render_template

from file_utils import append_row_to_spreadsheet
from file_utils import file_suffix
from file_utils import read_data_from_spreadsheet

app = Flask(__name__)

# root urls
localhost_root_url = 'http://127.0.0.1:5000/'
heroku_root_url = "https://concrete-flowchart.herokuapp.com/"


@app.route('/')
def root():
    return render_template('index.html',
                           full_summary_url=localhost_root_url + 'download_full',
                           short_summary_url=localhost_root_url + 'download_short')


@app.route('/download_full')
def download_full():
    path = pathlib.Path.cwd().joinpath('res')
    return send_from_directory(directory=path,
                               filename='report-' + file_suffix[1] + '-test.xlsx', as_attachment=True,
                               mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/download_short')
def download_short():
    path = pathlib.Path.cwd().joinpath('res')
    return send_from_directory(directory=path,
                               filename='report-' + file_suffix[0] + '-test.xlsx', as_attachment=True,
                               mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/short_summary', methods=['GET'])
def update_short_summary():
    args = request.args.to_dict()
    # print("Received request params: " + str(args))

    # append the new data to the previous spreadsheet
    result = append_row_to_spreadsheet(log=args, file_detail=file_suffix[0])
    if result:
        print("Short summary saved")

    # debug
    updated_data = read_data_from_spreadsheet('report-short-test.xlsx')
    return jsonify(updated_data.to_dict())


@app.route('/full_summary', methods=['GET'])
def update_full_summary():
    args = request.args.to_dict()
    # print("Received request params: " + str(args))

    # append the new data to the previous spreadsheet
    result = append_row_to_spreadsheet(log=args, file_detail=file_suffix[1])
    if result:
        print("Full summary saved")

    # debug
    updated_data = read_data_from_spreadsheet('report-full-test.xlsx')
    return jsonify(updated_data.to_dict())


if __name__ == '__main__':
    app.run()
