import os
from pathlib import Path
import sys
import dj_database_url

print("=" * 60)
print("Loading Django settings...")
print(f"Current directory: {os.getcwd()}")
print(f"Environment variables: RAILWAY={os.environ.get('RAILWAY', 'Not set')}")
print("=" * 60)

BASE_DIR = Path(__file__).resolve().parent.parent

SECRET_KEY = os.environ.get('SECRET_KEY', 'django-insecure-your-secret-key-here')

DEBUG = os.environ.get('DEBUG', 'False') == 'True'
print(f"DEBUG mode: {DEBUG}")

ALLOWED_HOSTS = [
    'localhost',
    '127.0.0.1',
    'newl-africa-production.up.railway.app',
    '.railway.app',
    os.environ.get('ALLOWED_HOSTS', '*'),
]

CSRF_TRUSTED_ORIGINS = [
    'https://newl-africa-production.up.railway.app',
    'https://*.railway.app',
]

# Add APPEND_SLASH setting to prevent redirect issues
APPEND_SLASH = True

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'core',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

ROOT_URLCONF = 'certificate_generator.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [BASE_DIR / 'templates'],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.debug',
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    },
]

WSGI_APPLICATION = 'certificate_generator.wsgi.application'

# Database Configuration
IS_PRODUCTION = os.environ.get('RAILWAY_ENVIRONMENT', False) or os.environ.get('RAILWAY_SERVICE_ID', False) or os.environ.get('RAILWAY', False)

print(f"IS_PRODUCTION: {IS_PRODUCTION}")

if IS_PRODUCTION:
    print("PRODUCTION MODE - Using PostgreSQL")
    DATABASE_URL = os.environ.get('DATABASE_URL')
    if DATABASE_URL:
        print("✅ DATABASE_URL found in environment")
    else:
        print("⚠️ DATABASE_URL not in environment, using hardcoded URL")
        DATABASE_URL = 'postgresql://postgres:LtgXUeszLiwmJMHIGIMmdKeiYVrvJiXZ@acela.proxy.rlwy.net:27018/railway'
    
    try:
        print(f"Connecting to database...")
        DATABASES = {
            'default': dj_database_url.config(default=DATABASE_URL, conn_max_age=600, ssl_require=True)
        }
        print("✅ PostgreSQL configuration successful")
    except Exception as e:
        print(f"❌ Database connection error: {e}")
        print("Falling back to SQLite...")
        DATABASES = {
            'default': {
                'ENGINE': 'django.db.backends.sqlite3',
                'NAME': BASE_DIR / 'db.sqlite3',
            }
        }
else:
    print("DEVELOPMENT MODE - Using SQLite")
    DATABASES = {
        'default': {
            'ENGINE': 'django.db.backends.sqlite3',
            'NAME': BASE_DIR / 'db.sqlite3',
        }
    }

print(f"Database ENGINE: {DATABASES['default']['ENGINE']}")
print("=" * 60)

AUTH_PASSWORD_VALIDATORS = [
    {'NAME': 'django.contrib.auth.password_validation.UserAttributeSimilarityValidator'},
    {'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator'},
    {'NAME': 'django.contrib.auth.password_validation.CommonPasswordValidator'},
    {'NAME': 'django.contrib.auth.password_validation.NumericPasswordValidator'},
]

LANGUAGE_CODE = 'en-us'
TIME_ZONE = 'Africa/Nairobi'
USE_I18N = True
USE_TZ = True

STATIC_URL = 'static/'
STATIC_ROOT = BASE_DIR / 'staticfiles'
STATICFILES_DIRS = [BASE_DIR / 'static']

MEDIA_URL = '/media/'
MEDIA_ROOT = BASE_DIR / 'media'

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

LOGIN_URL = 'login'
LOGIN_REDIRECT_URL = 'home'
LOGOUT_REDIRECT_URL = 'login'

DATA_UPLOAD_MAX_MEMORY_SIZE = 52428800
FILE_UPLOAD_MAX_MEMORY_SIZE = 52428800

if IS_PRODUCTION:
    SECURE_SSL_REDIRECT = True
    SESSION_COOKIE_SECURE = True
    CSRF_COOKIE_SECURE = True
    SECURE_BROWSER_XSS_FILTER = True
    SECURE_CONTENT_TYPE_NOSNIFF = True

print("✅ Settings loading complete!")
print("=" * 60)