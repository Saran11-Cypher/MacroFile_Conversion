from django.urls import path
from .views import user_login, user_logout, dashboard, user_signup, forgot_password, verify_otp, reset_password,admin_dashboard, make_admin, delete_user, upload, process_excel


urlpatterns = [
    path('', user_login, name='login'),  # Default page
    path('logout/', user_logout, name='logout'),
    path('dashboard/', dashboard, name='dashboard'),
    path('signup/', user_signup, name='signup'),
    path('forgot-password/', forgot_password, name='forgot_password'),
    path('verify-otp/', verify_otp, name='verify_otp'),
    path('reset-password/', reset_password, name='reset_password'),
    path('admin_dashboard/', admin_dashboard, name='admin_dashboard'),
    path('make_admin/<int:user_id>/', make_admin, name='make_admin'),
    path('delete_user/<int:user_id>/', delete_user, name='delete_user'),
    path('upload/', upload, name='upload'),
    path('process_excel/', process_excel, name='process_excel'),
]