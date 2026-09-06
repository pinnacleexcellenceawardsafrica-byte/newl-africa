"""
WSGI config for certificate_generator project.
"""

import os
import sys
import traceback

from django.core.wsgi import get_wsgi_application

try:
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'certificate_generator.settings')
    application = get_wsgi_application()
except Exception as e:
    print(f"Error loading WSGI application: {e}")
    traceback.print_exc()
    sys.exit(1)