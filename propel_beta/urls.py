from django.urls import path
from .views import chat, chat2


urlpatterns = [
    path('chat/', chat, name='chat'),
    path('chat2/', chat2, name='chat2'),
]
