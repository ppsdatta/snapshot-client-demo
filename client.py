import os
import sys
import mimetypes
import requests

SERVER_URL = 'http://localhost:8080/snapshot-server/'

TARGET_PDF = 'output.pdf'

MIME_MAPS = {
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'msexcel',
    'application/vnd.ms-excel': 'msexcel',
    'application/msword': 'msword',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'msword',
    'application/vnd.ms-powerpoint': 'msppt',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'msppt'
}


def read_file(file_name : str) -> bytes:
    """
    Reads the file from a given file name and returns the binary content of it.
    :param file_name: String absolute path of the file
    :return: bytearray of binary file content
    """

    with open(file_name, mode='rb') as f:
        file_content = f.read()

    return file_content


def send_request(url : str, file_content : bytes) -> bytes:
    """
    Sends the binary content of a file as request body to URL
    :param url : the url for the server
    :param file_content: bytes
    :return: bytes The binary response from server
    """

    payload = {'input': file_content}
    req = requests.post(url, data=payload)
    return req.content


def write_file(file_name : str, binary_content : bytes) -> None:
    """
    Writes the binary output from server into a pdf file
    :param file_name: String file name to write to.
    :param binary_content: content received from server
    :return: None
    """

    with open(file_name, 'wb') as f:
        f.write(binary_content)


def get_mmime_map_value(file_name : str) -> str:
    """
    Get the appropriate URL extension from mime type of a file
    :param file_name: String file name
    :return: Either a String mime map value - either of (msword, msexcel or msppt) or None
    """

    global MIME_MAPS

    file_base_name = os.path.basename(file_name)
    guess_type = mimetypes.guess_type(file_base_name)

    if guess_type[0] and (guess_type[0] in MIME_MAPS):
        return MIME_MAPS[guess_type[0]]
    else:
        return None


def generate_snapshot(file_name : str) -> None:
    """
    Main function to generate a pdf snapshot from a file
    :param file_name: String name of the file
    :return: None
    """

    global SERVER_URL
    global TARGET_PDF

    if not os.path.exists(file_name):
        raise ValueError('File {} could not be found'.format(file_name))

    mime_type = get_mmime_map_value(file_name)

    if not mime_type:
        raise ValueError('This type of file in not supported for snapshot generation')

    server_url = SERVER_URL + mime_type
    file_content = read_file(file_name)
    response_binary = send_request(server_url, file_content)
    write_file(TARGET_PDF, response_binary)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        sys.stderr.write('Please provide a file name!\n')
        sys.exit(2)

    generate_snapshot(sys.argv[1])

    print('trying to open {} automatically...'.format(TARGET_PDF))
    os.system('open {}'.format(TARGET_PDF))
    print('check {} for output'.format(TARGET_PDF))



