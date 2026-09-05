"""
WSGI config for pinnacle_awards project.
"""

import os

from django.core.wsgi import get_wsgi_application

# Change from 'mwasa.settings' to 'pinnacle_awards.settings'
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'pinnacle_awards.settings')

application = get_wsgi_application()