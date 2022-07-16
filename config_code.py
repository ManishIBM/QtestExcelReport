import json


class FileOperations:

    def __init__(self):

        self.download_file_path = ''

    def get_json_data(self, json_file_path):
        try:
            with open(json_file_path) as json_data:
                return json.load(json_data)
        except IOError as e:
            print("I/O error({0}): {1}".format(e.errno, e.strerror))
        except Exception as e:
            print('exception occured in %s having value %s',
                  __name__, str(e))
        return None

    def write_json_data(self, json_data, json_file_path):
        try:
            with open(json_file_path, "w") as json_file:
                json.dump(json_data, json_file)
                return True
        except IOError as e:
            print("I/O error({0}): {1}".format(e.errno, e.strerror))
        except Exception as e:
            print('exception occured in %s having value %s',
                  __name__, str(e))
        return False
