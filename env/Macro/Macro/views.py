from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.contrib.auth.models import User
from django.contrib import messages
from django.contrib.auth.decorators import login_required
import random
from django.core.mail import send_mail
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.contrib.auth.decorators import user_passes_test
import os
from django.shortcuts import render
from django.core.files.storage import default_storage
from openpyxl import Workbook, load_workbook
from .forms import FileUploadForm
from django.shortcuts import render
from django.core.files.storage import FileSystemStorage
from .models import GeneratedFile

# Function to check if user is admin
def is_admin(user):
    return user.is_superuser  # Only allow superusers

@login_required
@user_passes_test(is_admin)
def admin_dashboard(request):
    users = User.objects.all()  # Admin can see all users
    return render(request, 'admin_dash.html', {'users': users})

# Make a user an admin
@login_required
@user_passes_test(is_admin)
def make_admin(request, user_id):
    user = User.objects.get(id=user_id)
    user.is_superuser = True
    user.is_staff = True
    user.save()
    messages.success(request, f"{user.username} is now an admin!")
    return redirect('admin_dash')

# Delete a user
@login_required
@user_passes_test(is_admin)
def delete_user(request, user_id):
    user = User.objects.get(id=user_id)
    if not user.is_superuser:  # Prevent deletion of admins
        user.delete()
        messages.success(request, "User deleted successfully!")
    else:
        messages.error(request, "Cannot delete an admin user!")
    return redirect('admin_dash')

def user_login(request):
    if request.method == 'POST':
        username = request.POST['username']
        password = request.POST['password']

        user = authenticate(request, username=username, password=password)

        if user is not None:
            login(request, user)

            # Check if user is admin
            if user.is_staff or user.is_superuser:
                return redirect('/admin/')  # Redirect to Django admin panel

            return redirect('dashboard')  # Redirect to dashboard instead of 'home'

        else:
            messages.error(request, "Invalid username or password.")
            return redirect('login')

    return render(request, 'login.html')

def user_logout(request):
    logout(request)
    return redirect('login')  # Redirect to login page after logout

@login_required
def dashboard(request):
    generated_files = GeneratedFile.objects.all()
    return render(request, "dashboard.html", {"generated_files": generated_files})


def user_signup(request):
    if request.method == 'POST':
        username = request.POST['username']
        email = request.POST['email']
        password = request.POST['password']
        confirm_password = request.POST['confirm_password']

        if password != confirm_password:
            return render(request, 'signup.html', {'error': 'Passwords do not match'})

        if User.objects.filter(username=username).exists():
            return render(request, 'signup.html', {'error': 'Username already taken'})

        if User.objects.filter(email=email).exists():
            return render(request, 'signup.html', {'error': 'Email already registered'})

        user = User.objects.create_user(username=username, email=email, password=password)
        user.save()

        messages.success(request, "Signup successful! Please log in.")
        return redirect('login')  # Redirect to the login page

    return render(request, 'signup.html')

def forgot_password(request):
    if request.method == "POST":
        email = request.POST.get("email")

        if email:
            try:
                otp = str(random.randint(100000, 999999))  # Generate 6-digit OTP
                
                # Store OTP in session
                request.session["otp"] = otp
                request.session["email"] = email
                print("Stored OTP:", otp)  # Debugging

                # Send OTP via email
                send_mail(
                    "Password Reset OTP",
                    f"Your OTP for password reset is {otp}",
                    "your-email@gmail.com",  # Replace with your valid sender email
                    [email],
                    fail_silently=False,
                )

                return redirect("verify_otp")  # Check if URL exists in urls.py

            except Exception as e:
                print("Email sending error:", str(e))  # Debugging
                return render(request, "forgot_password.html", {"error": f"Error: {str(e)}"})

    return render(request, "forgot_password.html")

# def forgot_password(request):
#     if request.method == 'POST':
#         print("✅ POST request received.")  # Debug
#         email = request.POST['email']
#         print(f"✅ Email entered: {email}")  # Debug
        
#         try:
#             user = User.objects.get(email=email)
#             print("✅ User found in database.")  # Debug
            
#             otp = random.randint(100000, 999999)
#             request.session['otp'] = otp
#             request.session['email'] = email
#             print(f"✅ Generated OTP: {otp}")  # Debug

#             # Debug email sending
#             send_mail(
#                 'Your OTP Code',
#                 f'Your OTP code is {otp}',
#                 'saransuresh01s@gmail.com',
#                 [email],
#                 fail_silently=False,
#             )
#             print("✅ Email sent successfully.")  # Debug
            
#             return redirect('verify_otp')

#         except User.DoesNotExist:
#             print("❌ User not found.")  # Debug
#             return render(request, 'forgot_password.html', {'error': 'Email not found.'})
#     print("❌ Request was not POST")
#     return render(request, 'forgot_password.html')

def verify_otp(request):
    if request.method == "POST":
        entered_otp = request.POST.get("otp")
        stored_otp = request.session.get("otp")

        if entered_otp == stored_otp:
            del request.session["otp"]  # Remove OTP after successful verification
            return redirect('reset_password')  # Redirect to reset password page
        else:
            messages.error(request, "Invalid OTP. Please try again.")

    return render(request, "verify_otp.html")

@csrf_exempt
def send_otp(request):
    if request.method == "POST":
        print("OTP function called")
        email = request.POST.get("email")
        print("Email received:", email)

        if not email:
            return JsonResponse({"error": "Email is required"}, status=400)

        # Generate a 6-digit OTP
        otp = str(random.randint(100000, 999999))

        # Store the OTP in session
        request.session["otp"] = otp
        request.session["email"] = email
        print("Stored OTP in session:", otp)  # Debugging

        # Send OTP via email
        try:
            send_mail(
                "Your OTP Code",
                f"Your OTP is {otp}. Do not share it with anyone.",
                "saransuresh01s@gmail.com",
                [email],
                fail_silently=False,
            )
            print("OTP sent successfully")  # Debugging
            return JsonResponse({"message": "OTP sent successfully"})
        except Exception as e:
            print("Error sending email:", str(e))  # Debugging
            return JsonResponse({"error": "Failed to send OTP"}, status=500)

    return JsonResponse({"error": "Invalid request"}, status=400)

def reset_password(request):
    if request.method == "POST":
        password = request.POST['password']
        confirm_password = request.POST['confirm_password']

        # Check if passwords match
        if password != confirm_password:
            messages.error(request, "Passwords do not match!")
            return render(request, 'reset_password.html')

        # Get email from session
        email = request.session.get('email')
        if email:
            try:
                user = User.objects.get(email=email)
                
                # Reset the password
                user.set_password(password)
                user.save()

                # Clear the session to prevent unintended access to the reset flow
                del request.session['email']  # Optional: clear session email after reset

                messages.success(request, "Password reset successful! You can now log in.")
                return redirect('login')  # Redirect to login page after successful reset
            except User.DoesNotExist:
                messages.error(request, "User with this email does not exist.")
                return redirect('forgot_password')  # Redirect back to forgot password if no user found

    # If request method is not POST, render the reset password page
    return render(request, 'reset_password.html')

def process_excel(request):
    if request.method == "POST":
        form = FileUploadForm(request.POST, request.FILES)
        if form.is_valid():
            excel_file = request.FILES["excel_file"]
            file_path = default_storage.save("uploaded_files/" + excel_file.name, excel_file)
            full_path = os.path.join(default_storage.location, file_path)

            wb = load_workbook(full_path)
            sheet = wb.active  
            last_row = sheet.max_row

            output_dir = os.path.join(default_storage.location, "generated_files")
            os.makedirs(output_dir, exist_ok=True)

            file_links = []  # Store file paths for UI display

            for i in range(1, last_row + 1):
                cfname = sheet.cell(row=i, column=21).value  

                if not cfname:
                    continue  

                subfolder = os.path.join(output_dir, cfname)
                os.makedirs(subfolder, exist_ok=True)
                new_file_path = os.path.join(subfolder, f"UFT{cfname}.xlsx")

                new_wb = Workbook()
                new_ws = new_wb.active

                headers = ["Run", "Value", "FileName", "Error", "Status", "Export Date"]
                new_ws.append(headers)

                new_wb.save(new_file_path)

                # Save to database for UI display
                generated_file = GeneratedFile(file_name=f"UFT{cfname}.xlsx", file_path=new_file_path)
                generated_file.save()

                file_links.append(new_file_path)

            return render(request, "dashboard.html", {"file_links": file_links})

    return render(request, "dashboard.html", {"form": form})

def upload(request):
    if request.method == 'POST' and request.FILES.get('file'):
        file = request.FILES['file']
        fs = FileSystemStorage()
        filename = fs.save(file.name, file)
        uploaded_file_url = fs.url(filename)
        return render(request, 'dashboard.html', {'uploaded_file_url': uploaded_file_url})
    return render(request, 'dashboard.html')

