// sign up form
document.addEventListener('DOMContentLoaded', function() {
    const showSignup = document.getElementById('showSignup');
    const signupModal = document.getElementById('signupModal');
    const closeModal = document.querySelector('.close');
    const signupForm = document.getElementById('signupForm');
    const signupYear = document.getElementById('signupYear');
    const signupMajorGroup = document.getElementById('signupMajorGroup');

    if (showSignup) {
        showSignup.addEventListener('click', function(e) {
            e.preventDefault();
            signupModal.style.display = 'block';
        });
    }

    if (closeModal) {
        closeModal.addEventListener('click', function() {
            signupModal.style.display = 'none';
        });
    }

    // show lang major kung 3rd or 4th year
    if (signupYear && signupMajorGroup) {
        signupYear.addEventListener('change', function() {
            if (this.value === '3' || this.value === '4') {
                signupMajorGroup.style.display = 'block';
            } else {
                signupMajorGroup.style.display = 'none';
            }
        });
    }

    // signup
    if (signupForm) {
        signupForm.addEventListener('submit', function(e) {
            e.preventDefault();
            
            const formData = {
                studentNumber: document.getElementById('signupStudentNumber').value,
                name: document.getElementById('signupName').value,
                email: document.getElementById('signupEmail').value,
                password: document.getElementById('signupPassword').value,
                year: document.getElementById('signupYear').value,
                section: document.getElementById('signupSection').value,
                major: document.getElementById('signupMajor').value,
                sig: document.getElementById('signupSIG').value
            };

            //nag validate nong format ng student number
            const studentNumberRegex = /^02\d{2}-\d{4}$/;
            if (!studentNumberRegex.test(formData.studentNumber)) {
                alert('Please enter a valid student number (format: 02**-****)');
                return;
            }

            // password checker
            if (formData.password !== document.getElementById('signupConfirmPassword').value) {
                alert('Passwords do not match!');
                return;
            }

            fetch('/signup', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(formData)
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    alert('Application submitted! Please wait for Officer approval.');
                    signupModal.style.display = 'none';
                    signupForm.reset();
                } else {
                    alert(data.message);
                }
            })
            .catch(error => {
                console.error('Error:', error);
                alert('An error occurred during signup.');
            });
        });
    }
});