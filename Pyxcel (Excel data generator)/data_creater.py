import faker
from faker import *
# Initialize Faker with Indian locale
fake = Faker('en_IN')

# Generate fake data
name = fake.name() # Generate a random name
address = fake.address() # Generate a random address
email = fake.email() # Generate a random email address
phone_number = fake.phone_number() # Generate a random phone number
date_of_birth = fake.date_of_birth(minimum_age=18, maximum_age=65)  # Generate realistic dates of birth
job = fake.job() # Generate a random job title
company = fake.company() # Generate a random company name

