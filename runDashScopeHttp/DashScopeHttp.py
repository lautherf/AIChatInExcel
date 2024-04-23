from flask import Flask, request, jsonify
import random
from http import HTTPStatus
import dashscope
from dashscope import Generation
import requests
import json

dashscope.api_key="sk-xxxxxxxxxx"

def test():
    url = 'http://localhost:5000/v1/chat/completions'
    headers = {'Content-Type': 'application/json'}
    data = {
        'model': 'qwen1.5-72b-chat',
        'temperature': 0.7,
        'n': 3,
        'messages': [
            {'role': 'system', 'content': 'You are a helpful assistant.'},
            {'role': 'user', 'content': '如何做西红柿炒鸡蛋？'}
        ]
    }

    json_data = json.dumps(data)

    try:
        response = requests.post(url, headers=headers, data=json_data)

        if response.status_code == 200:
            content = response.json()
            print('Response:')
            print(content)
        else:
            print('Error:', response.status_code)
    except requests.exceptions.RequestException as e:
        print('Error:', e)

def DashScopeCall(model,messages,**kwargs):
    response = Generation.call(
        model = model,
        messages=messages,
        # set the random seed, optional, default to 1234 if not set
        seed=random.randint(1, 10000),
        result_format='message',  # set the result to be "message"  format.
        temperature = kwargs.get('temperature', 0.7),
    )
    if response.status_code == HTTPStatus.OK:
        print('DashScopeCall response = ')
        print(response)
    else:
        print('Request id: %s, Status code: %s, error code: %s, error message: %s' % (
            response.request_id, response.status_code,
            response.code, response.message
        ))
    return response


app = Flask(__name__)

@app.route('/v1/chat/completions', methods=['POST'])
def generate():
    data = request.get_json()
    model = data.get('model', 'qwen1.5-72b-chat')
    temperature = data.get('temperature', 0.7)
    responseNum = data.get('n', 1)
    messages = data.get('messages', '')

    request_id_list = ''
    choicesArray = []
    prompt_tokens_total = completion_tokens_total = total_tokens_total = 0

    for i in range(responseNum):
        try:
            response = DashScopeCall(model, messages, temperature = temperature);
        except Exception as e:
            return jsonify({'error': str(e)}), HTTPStatus.INTERNAL_SERVER_ERROR

        if response.status_code == HTTPStatus.OK:
            request_id_list = request_id_list + response.request_id
            message = response.output.choices[0].message
            finish_reason = response.output.choices[0].finish_reason
            choicesItem = {
                'index' : i ,
                'message': message ,
                'finish_reason': finish_reason ,
            }
            choicesArray.append(choicesItem)
            prompt_tokens_total += response.usage.input_tokens
            completion_tokens_total += response.usage.output_tokens
            total_tokens_total += response.usage.total_tokens

    formatResponse = {
        'id': request_id_list,
        'model': model,
        'choices': choicesArray,
        'usage': {
            'prompt_tokens': prompt_tokens_total,
            'completion_tokens': completion_tokens_total,
            'total_tokens': total_tokens_total,
        },
    }

    print('DashScopeCall formatResponse = ')
    print(formatResponse)

    return jsonify(formatResponse), HTTPStatus.OK

if __name__ == '__main__':
    app.run(debug=True)