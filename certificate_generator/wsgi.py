"""
WSGI config for certificate_generator project.
"""

import os
import sys
import traceback

# Print Python path for debugging
print("=" * 60)
print("Starting WSGI application...")
print(f"Python version: {sys.version}")
print(f"Current directory: {os.getcwd()}")
print(f"Files in current dir: {os.listdir('.')}")
print("=" * 60)

try:
    # Set settings module
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'certificate_generator.settings')
    print("✅ DJANGO_SETTINGS_MODULE set to certificate_generator.settings")
    
    # Import settings to test
    from django.conf import settings
    print(f"✅ Settings loaded. DEBUG={settings.DEBUG}")
    print(f"✅ Database engine: {settings.DATABASES['default']['ENGINE']}")
    print(f"✅ Database name: {settings.DATABASES['default']['NAME']}")
    
    # Test database connection
    from django.db import connection
    with connection.cursor() as cursor:
        cursor.execute("SELECT 1")
        print("✅ Database connection successful!")
    
    # Now load the WSGI application
    from django.core.wsgi import get_wsgi_application
    application = get_wsgi_application()
    print("✅ WSGI application loaded successfully!")
    print("=" * 60)
    
except Exception as e:
    print(f"❌ ERROR loading WSGI application: {e}")
    print("=" * 60)
    traceback.print_exc()
    print("=" * 60)
    # Don't exit - let it fail with proper error
    raise