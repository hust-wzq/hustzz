import json
from tencentcloud.common import credential
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.asr.v20190614 import asr_client, models
import base64
import easygui as g


class STT:
    def __init__(self):
        cred = credential.Credential("AKIDXu0pOhjUyB7miTzFG2Yopa7OPAgwve9O", "0lHXXHlzwc08zZsvmOlskMioeEWXJn0w")
        httpProfile = HttpProfile()
        httpProfile.endpoint = "asr.tencentcloudapi.com"
        clientProfile = ClientProfile()
        clientProfile.httpProfile = httpProfile
        self.client = asr_client.AsrClient(cred, "", clientProfile)
        self.req = models.CreateRecTaskRequest()

    def trans(self, bs64):
        try:
            params = {
                "EngineModelType": "16k_zh",
                "ChannelNum": 1,
                "SpeakerDiarization": 1,
                "SpeakerNumber": 0,
                "ResTextFormat": 0,
                "SourceType": 0,
                "Data": "111",
                "FilterModal": 2
            }
            self.req.from_json_string(json.dumps(params))

            resp = self.client.CreateRecTask(self.req)
            print(resp.to_json_string())

        except TencentCloudSDKException as err:
            print(err)


class CryptFile:
    def __init__(self, file):
        f = open(file, 'rb')
        filestr = f.read()
        self.encodestr = base64.b64encode(filestr)


if __name__ == '__main__':
    file = g.fileopenbox()
    crapp = CryptFile(file)
    bs64str = crapp.encodestr
    app = STT()
    res = app.trans(bs64str)
