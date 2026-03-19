import requests


def get_file_from_vm():
    resp = requests.post(
        url="http://[fdbd:dc02:fd:4::7e]:5000/file",
        data={"file_path": "C:\\Users\\User\\OSWorld\\desktop_env\\server.zip"},
    )

    with open("server.zip", "wb") as f:
        f.write(resp.content)


def get_file_from_vm_ubuntu():

    resp = requests.post(
        url="http://[fdbd:dc02:fd:1:1::120]:5000/file",
        data={"file_path": "/home/user/server.tar.gz"},
    )

    with open("server.tar.gz", "wb") as f:
        f.write(resp.content)


if __name__ == "__main__":
    get_file_from_vm_ubuntu()
