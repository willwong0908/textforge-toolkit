"""Isolated, visible-browser QA host; only model responses are simulated."""
from pathlib import Path
import asyncio
import json
import sys
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
QA = ROOT / 'output' / 'playwright' / 'review-learning'
QA.mkdir(parents=True, exist_ok=True)

from term_extractor_app import storage, telemetry
patch.object(storage, 'get_app_root', return_value=QA).start()
patch.object(telemetry, 'track_event').start()
store = storage.SettingsStore(storage.get_app_paths())
settings = store.load()
settings.provider_settings[settings.provider_name].api_key = 'qa-dummy-key'
settings.provider_settings[settings.provider_name].model = 'qa-simulated-model'
store.save(settings)

from term_extractor_app.ai_review import database, session_store, review_service, output_service, shared_provider, workspace_service
from term_extractor_app.providers import OpenAICompatibleAdapter
from term_extractor_app.models import LLMResponse
from test_ai_review_conversation_api import _workspace_response


async def simulated_response(self, request, attempt=1):
    await asyncio.sleep(.2)
    package = request.metadata.get('package', [])
    results = [{'id':i['id'], 'has_issue':True, 'issue_type':'术语',
                'issue':'译文未逐字使用推荐术语（模拟误报）', 'suggestion':'You promote.'} for i in package]
    return LLMResponse(task_id=request.task_id, task_type=request.task_type, content=json.dumps({'items':results}),
                       provider='QA', model='qa-simulated-model', success=True, attempts=attempt, latency_ms=200)


def simulated_learning(task_id, messages, **kwargs):
    if task_id.startswith('memory_'):
        payload=json.loads(messages[-1]['content'])
        old=payload.get('old_rules', [])
        return json.dumps({'triggered_ids':[old[0]['id']] if old else [],
                           'new_rules':[] if old else ['英语中应结合词性与语境使用推荐术语；自然准确的名词表达不机械替换为动词。']})
    return '这是一条模拟追问回复：应结合语境判断。'


def simulated_workspace(messages, on_delta=None):
    response = _workspace_response(messages)
    response = response.replace('"scope": "table"', '"scope": "Terms QA"')
    if on_delta:
        on_delta(response, 'content')
    return response


patch.object(OpenAICompatibleAdapter, 'send_prompt', simulated_response).start()
patch.object(shared_provider, 'followup_chat', simulated_learning).start()
patch.object(review_service, 'followup_chat', simulated_learning).start()
patch.object(workspace_service, 'workspace_chat', simulated_workspace).start()

database.init_db()
from term_extractor_app.web_app import app
from openpyxl import Workbook
book=Workbook(); sheet=book.active; sheet.title='Terms QA'
sheet.append(['English', 'Chinese', 'Japanese'])
sheet.append(['You gain promotions.', '你们升星。', '昇格します。'])
sheet.append(['I bring gifts.', '我发礼物。', '贈り物です。'])
book.save(QA / 'review-input.xlsx'); book.close()
print('QA_ARTIFACTS=' + str(QA), flush=True)

if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host='127.0.0.1', port=8876, access_log=False)
