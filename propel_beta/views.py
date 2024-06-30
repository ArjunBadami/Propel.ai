from django.shortcuts import render
from django.http import JsonResponse
#from .models import Agent
import autogen
from django.http import JsonResponse
from django.views.decorators.http import require_http_methods
from django.views.decorators.csrf import csrf_exempt
import json
from .agents import llm_config
from .agents import call_chat
from .agents import getMonthlyReport, getcpm, excel_reader, excel_writer, extract_text_from_pdf, retrieve_respond, getPaymentMilestones
import traceback

key = "temp_key"
chat_manager = {}
# manager = autogen.GroupChatManager(groupchat=groupchat, llm_config=llm_config) 
def chat2(request):
    #items = list(Agent.objects.values('name', 'description'))  # Fetching data from database
    #return request.method
    return JsonResponse({'response': request.method})


@csrf_exempt  # Use this decorator to exempt the view from CSRF verification.
#@require_http_methods(["POST"])  # This view will only accept POST requests.
def chat(request):
  # try:
    data = json.loads(request.body)
    prompt = "TERMINATE"
    if('input' in data):
      prompt = data['input']

    target_phrase = "terms of payment and invoicing"
    replacement_phrase = target_phrase.upper()
    if target_phrase in prompt:
        prompt = prompt.replace(target_phrase, replacement_phrase)
  
    '''
    manager = None
    if chat_manager.get(key, None) == None:
      manager = autogen.GroupChatManager(groupchat=groupchat, llm_config=llm_config) 
      chat_manager[key] = manager
    else:
      manager = chat_manager.get(key)
    
    message = boss.initiate_chat(
        manager,
        clear_history=False,
        #message=boss_aid.message_generator,
        message=prompt,
        #n_results=3,
    )
    '''
    message = call_chat(prompt)
    last_non_empty_content = None
    for item in reversed(message.chat_history):
        if item['content'].strip() != '' and item['content'].strip() != 'TERMINATE' and item['content'].strip() != 'Reply `TERMINATE` if the task is done.':
            last_non_empty_content = item['content']
        #else:
            break
    #messages = boss.chat_messages[manager] 
    #last_message = messages[len(messages) - 1]['content']

    function_call = last_non_empty_content#.replace('"', r'\"')
    function_name = function_call.split('(')[0]
    result = ''
    
    try:
      result = eval(function_call)
    except Exception as e:
      traceback.print_exc()
      print(e.__str__())
      result = "Error"  
    
    if result.endswith("TERMINATE"):
        result = result[:-len("TERMINATE")]

    return JsonResponse({"message": result})
    #return JsonResponse(json.loads(request.body))
  # except Exception as e:
  #   return JsonResponse({"message": e.args[0]})