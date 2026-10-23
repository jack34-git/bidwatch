from django.core.wsgi import get_wsgi_application
from django.urls import path

import app

urlpatterns = [
    path("", app.index),
    path("login/", app.login_view),
    path("logout/", app.logout_view),
    path("save/", app.save_settings),
    path("search/", app.search_now),
    path("excel/<str:run_id>/", app.download_excel),
]

application = get_wsgi_application()
