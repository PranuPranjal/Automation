import openpyxl
import random

# Load the previously generated names from the uploaded file to avoid duplication
previous_names = set()
img_path = "C:/Users/Lenovo/Downloads/compressedImages/"

# with open("new_indian_names3.csv", mode="r") as file:
#     reader = csv.reader(file)
#     next(reader)  # Skip the header
#     for row in reader:
#         previous_names.add(row[0])  # Add the name (first column) to the set

Plan = ["Yes", "No"]

Mode = ["Train", "Self-drive", "Air", "Self-drive", "Road", "Air", "Air", "Air", "Train", "Self-drive", "Air", "Road", "Self-drive", "Air", "Road", "Air", "Train", "Air", "Self-drive", "Air", "Self-drive", "Air", "Road", "Air", "Road", "Road", "Train", "Self-drive", "Self-drive", "Train", "Road", "Air", "Self-drive", "Road", "Train", "Road", "Road", "Self-drive", "Road", "Self-drive", "Self-drive", "Train", "Train", "Train", "Train", "Air", "Self-drive", "Train", "Road", "Self-drive", "Road", "Air", "Road", "Road", "Self-drive", "Air", "Road", "Self-drive", "Self-drive", "Road", "Train", "Air", "Self-drive", "Road", "Road", "Self-drive", "Road", "Road", "Self-drive", "Self-drive", "Road", "Train", "Road", "Road", "Self-drive", "Road", "Road", "Road", "Air", "Train", "Self-drive", "Train", "Air", "Air", "Train", "Train", "Train", "Self-drive", "Air", "Self-drive", "Train", "Train", "Air", "Air", "Self-drive", "Road", "Air", "Self-drive", "Train", "Air", "Air", "Air"]

Group =["With friends", "With friends", "Solo", "With family", "With friends", "Solo", "With family", "With family", "Solo", "With friends", "Solo", "With family", "Solo", "Solo", "With family", "Solo", "With family", "With family", "With friends", "With family", "Solo", "With friends", "With friends", "With family", "With friends", "With friends", "Solo", "With friends", "Solo", "With family", "With friends", "Solo", "With family", "With friends", "With family", "With family", "Solo", "Solo", "With family", "With family", "With family", "Solo", "With family", "With friends", "With friends", "Solo", "Solo", "With family", "With family", "With family", "Solo", "With family", "With family", "Solo", "With family", "Solo", "With friends", "With friends", "With friends", "Solo", "With family", "With friends", "Solo", "With family", "With friends", "Solo", "Solo", "With family", "With friends", "With family", "With friends", "With family", "Solo", "With family", "Solo", "Solo", "Solo", "Solo", "With family", "With family", "Solo", "Solo", "With friends", "Solo", "Solo", "With family", "With family", "Solo", "With friends", "With family", "Solo", "With family", "Solo", "Solo", "With friends", "Solo", "With family"]

Travel = [20000, 5000, 12500, 2500, 5000, 9100, 17800, 15000, 12500, 17800, 9100, 12500, 17800, 5000, 17800, 5000, 20000, 17800, 1600, 2500, 12500, 2500, 17800, 17800, 2500, 15000, 20000, 1600, 20000, 3900, 17800, 15000, 5000, 3900, 2500, 9100, 12500, 15000, 3900, 12500, 7200, 9100, 2500, 7200, 20000, 2500, 17800, 12500, 7200, 9100, 9100, 5000, 2500, 1600, 9100, 17800, 3900, 15000, 7200, 17800, 7200, 15000, 17800, 12500, 9100, 15000, 3900, 20000, 12500, 1600, 7200, 17800, 2500, 7200, 12500, 2500, 9100, 20000, 1600, 12500, 17800, 17800, 1600, 20000, 7200, 7200, 2500, 15000, 17800, 17800, 9100, 7200, 15000, 2500]


Food = [2500, 2500, 3900, 3900, 2500, 15000, 2500, 5000, 3900, 7200, 17800, 3900, 15000, 15000, 17800, 2500, 5000, 1600, 3900, 9100, 5000, 12500, 7200, 17800, 2500, 17800, 1600, 12500, 17800, 12500, 5000, 12500, 20000, 15000, 5000, 9100, 1600, 20000, 9100, 7200, 12500, 7200, 17800, 2500, 12500, 17800, 1600, 2500, 17800, 12500, 1600, 9100, 1600, 9100, 17800, 15000, 2500, 9100, 15000, 1600, 7200, 9100, 12500, 1600, 1600, 9100, 17800, 12500, 1600, 5000, 1600, 15000, 12500, 17800, 7200, 9100, 12500, 15000, 2500, 9100, 20000, 17800, 17800, 2500]

 
ReligiousItems = [17800, 9100, 2500, 5000, 2500, 2500, 5000, 7200, 17800, 7200, 15000, 2500, 7200, 1600, 1600, 15000, 15000, 3900, 9100, 15000, 17800, 15000, 17800, 9100, 2500, 15000, 17800, 20000, 12500, 9100, 15000, 7200, 1600, 20000, 9100, 9100, 7200, 2500, 12500, 15000, 9100, 9100, 20000, 2500, 15000, 12500, 1600, 17800, 3900, 7200, 7200, 7200, 9100, 20000, 12500, 9100, 15000, 1600, 20000, 12500, 17800, 2500, 9100, 2500, 1600, 9100, 1600, 17800, 2500, 15000, 2500, 20000, 1600, 15000, 15000, 17800, 1600, 20000, 1600, 2500]


Recreation = [15000, 2500, 3900, 2500, 12500, 7200, 2500, 3900, 7200, 9100, 12500, 2500, 7200, 3900, 3900, 5000, 3900, 17800, 1600, 2500, 5000, 5000, 20000, 17800, 17800, 2500, 12500, 17800, 12500, 12500, 7200, 20000, 12500, 5000, 17800, 9100, 12500, 9100, 17800, 9100, 7200, 7200, 12500, 12500, 12500, 2500, 20000, 15000, 2500, 12500, 7200, 17800, 9100, 17800, 12500, 17800, 7200, 2500, 15000, 12500, 17800, 15000, 2500, 9100, 1600, 9100, 7200, 2500, 9100, 2500, 1600, 1600, 17800, 17800, 17800, 12500, 1600, 15000, 7200, 20000, 17800, 1600, 12500]

Shopping = [9100, 3900, 17800, 15000, 2500, 15000, 9100, 15000, 17800, 7200, 15000, 9100, 17800, 9100, 20000, 3900, 9100, 15000, 7200, 9100, 20000, 17800, 17800, 2500, 1600, 9100, 1600, 17800, 1600, 3900, 12500, 20000, 1600, 17800, 20000, 17800, 12500, 2500, 9100, 2500, 7200, 2500, 1600, 7200, 2500, 12500, 7200, 20000, 7200, 15000, 3900, 20000, 9100, 1600, 7200, 17800, 1600, 3900, 7200, 15000, 9100, 12500, 2500, 1600, 3900, 5000, 7200, 15000, 5000, 20000, 20000, 17800, 20000, 7200, 5000, 17800, 12500, 17800, 20000, 15000, 15000, 3900, 2500, 12500, 9100, 20000, 3900, 15000, 17800, 20000, 12500, 17800, 20000, 20000, 20000, 12500, 3900, 3900, 3900, 17800, 7200]


Others = [9100, 3900, 17800, 15000, 2500, 15000, 9100, 15000, 17800, 7200, 15000, 9100, 17800, 9100, 20000, 3900, 9100, 15000, 7200, 9100, 20000, 17800, 17800, 2500, 1600, 9100, 1600, 17800, 1600, 3900, 12500, 20000, 1600, 17800, 20000, 17800, 12500, 2500, 9100, 2500, 7200, 2500, 1600, 7200, 2500, 12500, 7200, 20000, 7200, 15000, 3900, 20000, 9100, 1600, 7200, 17800, 1600, 3900, 7200, 15000, 9100, 12500, 2500, 1600, 3900, 5000, 7200, 15000, 5000, 20000, 20000, 17800, 20000, 7200, 5000, 17800, 12500, 17800, 20000, 15000, 15000, 3900, 2500, 12500, 9100, 20000, 3900, 15000, 17800, 20000, 12500, 17800, 20000, 20000, 20000, 12500, 3900, 3900, 3900, 17800, 7200]

Visits = [
    "Taj Mahal", "Jaipur City Palace", "Golden Temple", "Mysore Palace", "Qutub Minar",
    "Gateway of India", "Hawa Mahal", "Charminar", "Red Fort", "Victoria Memorial",
    "Sun Temple", "Mehrangarh Fort", "Amber Fort", "Ajanta Caves", "Ellora Caves",
    "Rann of Kutch", "Kumarakom Backwaters", "Meenakshi Temple", "Bandipur National Park",
    "Jim Corbett National Park", "Ranthambore National Park", "Kaziranga National Park",
    "Sundarbans", "Chittorgarh Fort", "Udaipur City Palace", "Lotus Temple", "Sanchi Stupa",
    "Brihadeeswarar Temple", "Jaisalmer Fort", "Spiti Valley", "Leh Palace", "Rishikesh",
    "Valley of Flowers", "Mahabalipuram", "Gwalior Fort", "Andaman Islands", "Shaniwar Wada",
    "Elephanta Caves", "Chandni Chowk", "Kanyakumari", "Darjeeling", "Chilka Lake",
    "Alleppey Backwaters", "Mount Abu", "Munnar", "Coorg", "Pondicherry", "Shimla",
    "Manali", "Kodaikanal", "Ramoji Film City", "Nandi Hills", "Auroville", "Tawang",
    "Hampi", "Varkala Beach", "Gokarna", "Chikmagalur", "Lakshadweep Islands",
    "Pangong Lake", "Tso Moriri Lake", "Dudhsagar Falls", "Baga Beach", "Vagator Beach",
    "Arambol Beach", "Kumarakom", "Wayanad", "Ravangla", "Gangtok", "Lachen",
    "Lachung", "Majuli Island", "Bomdila", "Ziro Valley", "Mawlynnong", "Cherrapunji",
    "Shillong", "Dawki", "Nohkalikai Falls", "Nathula Pass", "Rohtang Pass", "Solang Valley",
    "Kufri", "Dalhousie", "Khajjiar", "Bir Billing", "Mcleodganj", "Tirthan Valley",
    "Gir National Park", "Dwarka", "Somnath", "Saputara", "Modhera Sun Temple",
    "Lothal", "Bhuj", "Diu", "Mandvi Beach", "Mount Harriet National Park",
    "Chopta", "Hemkund Sahib", "Haridwar", "Badrinath", "Kedarnath", "Gangotri"
]

# Define new name lists for male and female
old_male_names = ["Prasad", "RamKrishna", "Harishankar", "Vijayshankar", "Vivekanand", "Radheshyam", 
                  "Baburao", "Babu", "Ramdas", "Tukaram", "Ramchandra", "Bhimrao", "Ramakant", "Prithviraj",
                    "Gangadhar", "Balkrishan", "Ramdev", "Vishwanath", "Ishwar", "Bhagwan", "Shankar"]

old_female_names = ["Meera", "Vasudha", "Savitri", "Gayatri", "Parvati", "Kaushalya", "Sumitra",
                  "Ganga", "Saraswati", "Jaya", "Rekha", "Sushma", "Nirmala", "Saroj", "Pushpa",
                   "Malti", "Hema", "Urmila"]

new_male_names = ["Aditya", "Prashant", "Yash", "Rohit", "Devansh", "Vikas", "Arjun", "Manish", "Shreyas", "Arpit",
                   "Rohan", "Harsh", "Neeraj", "Parth", "Dhruv", "Ravi", "Vivek", "Keshav", "Siddhant", "Mohit",
                   "Ishaan", "Arnav", "Kartik", "Dev", "Mihir", "Samar", "Aarav", "Kabir", "Vivaan", "Yuvraj", "Ritvik", 
                   "Shaurya", "Jatin", "Nirav", "Ved", "Harshit", "Aryan", "Lakshya", "Om", "Soham","Divyansh","Sushant",
                   "Abhimanyu","Mayank","Pranjal", "Aniket", "Amit", "Saurabh", "Gaurav", "Nikhil", "Shubham", "Aayush", 
                   "Rahul", "Varun", "Adarsh", "Sanjeev", "Tushar", "Akash", "Raj", "Vijay", "Ajay", "Ashish", "Uday",
                    "Naveen", "Anshul", "Siddharth", "Himanshu", "Krishna", "Sameer", "Vikram", "Kunal", "Suraj", "Tarun", 
                    "Anirudh", "Jay", "Tejas", "Abhinav", "Ayaan", "Rajesh", "Pankaj", "Ajit", "Deepak", "Shivam", "Manoj",
                     "Nitin", "Sumit", "Vatsal", "Bhavesh", "Hemant", "Lalit", "Raghav", "Nishant", "Mohan", "Sachin", "Rajan",
                     "Santosh", "Vinay", "Arvind", "Virat", "Rohit", "Ishant", "Ravi", "Arnav", "Anshul", "Mehul", "Swastik", "Prateek"]


new_female_names = ["Matki", "Ananya", "Sanya", "Ritu", "Nandini", "Ishita", "Tanvi", "Supriya", "Kavya", "Aditi", "Shruti",
                     "Radhika", "Vaishnavi", "Meera", "Tanisha", "Simran", "Amrita", "Isha", "Bhavya", "Janhvi", "Riya",
                     "Tanya", "Avantika", "Srishti", "Mahi", "Myra", "Lavanya", "Suhani", "Aarohi", "Tara", "Vidhi", "Siddhi",
                      "Mira", "Ruhi", "Dhriti", "Ira", "Ritika", "Esha", "Aarya", "Manvi", "Pihu","Sushmita","Muskan","Bhavana",
                      "Naina","Ayushi","Anushka", "Aditi", "Neha", "Anjali", "Sneha", "Priya", "Riya", "Pooja", "Megha", "Shreya",
                       "Kavya", "Sakshi", "Rashmi", "Tanvi", "Isha", "Nidhi", "Simran", "Tanya", "Ishita", "Swati", "Ananya", 
                       "Divya", "Sanya", "Avni", "Madhuri", "Sonali", "Aarohi", "Suhani", "Bhavya", "Vidhi", "Vaishnavi", "Kripa",
                        "Aarushi", "Jhanvi", "Esha", "Pragya", "Nisha", "Surbhi", "Ritika", "Payal", "Trisha",
                         "Garima", "Rupal", "Preeti", "Riddhi", "Mansi", "Pallavi", "Amrita", "Aishwarya", "Aparna", "Anushka"]

surnames = ["Sharma", "Patel", "Reddy", "Gupta", "Kumar", "Verma", "Nair", "Singh", "Mehta", "Desai", "Rao", "Chauhan", 
            "Das", "Iyer", "Bhat", "Agarwal", "Gandhi", "Jain", "Pillai", "Joshi", "Kapoor", "Shah", "Mishra", "Tripathi", 
            "Bhatt", "Sen", "Chakraborty", "Yadav", "Sinha", "Bose", "Chatterjee", "Malhotra", "Ghosh", "Pandey", "Rana", 
            "Naik","Murthy","Krishnan","Swamy","Selvam","Naidu","Chandrakar","Panigrahi","Pradhan","Dey","Dash","Gore","Kale",
            "Sanvale","Sonavane","Chakole","Talpade","Shikhre","Tiwari","Bajpai","Singh","Singhania","Galgotia","Rajpoot","Thapar",
            "Pahariya", "Chaudhary", "Bhagat", "Pawar", "Gokhale", "Chavan", "Rajput", "Rastogi", "Shetty", "Kulkarni", "Kashyap",
             "Dwivedi", "Pathak", "Garg", "Lal", "Thakur", "Roy", "Dasgupta", "Mukherjee", "Majumdar", "Kar", "Sarkar", "Chaudhari", 
             "Salvi", "Gadre", "Khatri", "Mohanty", "Saxena", "Shroff", "Deshmukh", "Gaikwad", "Raj", "Acharya", "Kohli", 
             "Sethi", "Puri", "Mahajan", "Bhonsle", "Bhadauria", "Ghoshal", "Kadam", "Sidhu", "Vohra", "Khurana", "Kalra"]
surnames_to_states = {
    'Sharma': ('Uttar Pradesh', ['Lucknow', 'Kanpur', 'Agra', 'Varanasi', 'Meerut', 'Bareilly', 'Allahabad', 'Moradabad', 'Aligarh', 'Ghaziabad']),
    'Patel': ('Gujarat', ['Ahmedabad', 'Surat', 'Vadodara', 'Rajkot', 'Bhavnagar', 'Jamnagar', 'Junagadh', 'Gandhinagar', 'Anand', 'Nadiad']),
    'Reddy': ('Telangana', ['Hyderabad', 'Warangal', 'Nizamabad', 'Khammam', 'Karimnagar', 'Ramagundam', 'Mahbubnagar', 'Mancherial', 'Siddipet', 'Miryalaguda']),
    'Gupta': ('Delhi', ['Delhi']),
    'Kumar': ('Bihar', ['Patna', 'Gaya', 'Bhagalpur', 'Muzaffarpur', 'Purnia', 'Darbhanga', 'Bihar Sharif', 'Arrah', 'Begusarai', 'Katihar']),
    'Verma': ('Uttar Pradesh', ['Lucknow', 'Kanpur', 'Agra', 'Varanasi', 'Meerut', 'Bareilly', 'Allahabad', 'Moradabad', 'Aligarh', 'Ghaziabad']),
    'Nair': ('Kerala', ['Thiruvananthapuram', 'Kochi', 'Kozhikode', 'Kollam', 'Thrissur', 'Alappuzha', 'Palakkad', 'Malappuram', 'Kannur', 'Kottayam']),
    'Singh': ('Punjab', ['Ludhiana', 'Amritsar', 'Jalandhar', 'Patiala', 'Bathinda', 'Mohali', 'Hoshiarpur', 'Batala', 'Moga', 'Abohar']),
    'Mehta': ('Rajasthan', ['Jaipur', 'Jodhpur', 'Kota', 'Ajmer', 'Udaipur', 'Bikaner', 'Alwar', 'Sikar', 'Bharatpur', 'Bhilwara']),
    'Desai': ('Gujarat', ['Ahmedabad', 'Surat', 'Vadodara', 'Rajkot', 'Bhavnagar', 'Jamnagar', 'Junagadh', 'Gandhinagar', 'Anand', 'Nadiad']),
    'Rao': ('Andhra Pradesh', ['Vijayawada', 'Visakhapatnam', 'Guntur', 'Nellore', 'Kurnool', 'Rajahmundry', 'Tirupati', 'Anantapur', 'Kadapa', 'Eluru']),
    'Chauhan': ('Uttar Pradesh', ['Lucknow', 'Kanpur', 'Agra', 'Varanasi', 'Meerut', 'Bareilly', 'Allahabad', 'Moradabad', 'Aligarh', 'Ghaziabad']),
    'Das': ('West Bengal', ['Kolkata', 'Asansol', 'Siliguri', 'Durgapur', 'Bardhaman', 'Malda', 'Kharagpur', 'Haldia', 'Raiganj', 'Baharampur']),
    'Iyer': ('Tamil Nadu', ['Chennai', 'Coimbatore', 'Madurai', 'Tiruchirappalli', 'Salem', 'Tirunelveli', 'Erode', 'Vellore', 'Tiruppur', 'Thoothukudi']),
    'Bhat': ('Jammu and Kashmir', ['Srinagar', 'Jammu', 'Anantnag', 'Baramulla', 'Kathua', 'Sopore', 'Udhampur', 'Pulwama', 'Poonch', 'Rajouri']),
    'Agarwal': ('Uttar Pradesh', ['Lucknow', 'Kanpur', 'Agra', 'Varanasi', 'Meerut', 'Bareilly', 'Allahabad', 'Moradabad', 'Aligarh', 'Ghaziabad']),
    'Gandhi': ('Gujarat', ['Ahmedabad', 'Surat', 'Vadodara', 'Rajkot', 'Bhavnagar', 'Jamnagar', 'Junagadh', 'Gandhinagar', 'Anand', 'Nadiad']),
    'Jain': ('Rajasthan', ['Jaipur', 'Jodhpur', 'Kota', 'Ajmer', 'Udaipur', 'Bikaner', 'Alwar', 'Sikar', 'Bharatpur', 'Bhilwara']),
    'Pillai': ('Kerala', ['Thiruvananthapuram', 'Kochi', 'Kozhikode', 'Kollam', 'Thrissur', 'Alappuzha', 'Palakkad', 'Malappuram', 'Kannur', 'Kottayam']),
    'Joshi': ('Maharashtra', ['Mumbai', 'Pune', 'Nagpur', 'Thane', 'Nashik', 'Aurangabad', 'Solapur', 'Amravati', 'Kolhapur', 'Jalgaon']),
    'Kapoor': ('Punjab', ['Ludhiana', 'Amritsar', 'Jalandhar', 'Patiala', 'Bathinda', 'Mohali', 'Hoshiarpur', 'Batala', 'Moga', 'Abohar']),
    'Shah': ('Gujarat', ['Ahmedabad', 'Surat', 'Vadodara', 'Rajkot', 'Bhavnagar', 'Jamnagar', 'Junagadh', 'Gandhinagar', 'Anand', 'Nadiad']),
    'Mishra': ('Uttarakhand', ['Dehradun', 'Haridwar', 'Roorkee', 'Haldwani', 'Rudrapur', 'Rishikesh', 'Kashipur', 'Kotdwar', 'Pithoragarh', 'Nainital']),
    'Tripathi': ('Uttar Pradesh', ['Lucknow', 'Kanpur', 'Agra', 'Varanasi', 'Meerut', 'Bareilly', 'Moradabad', 'Aligarh', 'Ghaziabad']),
    'Bhatt': ('Uttarakhand', ['Dehradun', 'Haridwar', 'Roorkee', 'Haldwani', 'Rudrapur', 'Rishikesh', 'Kashipur', 'Kotdwar', 'Pithoragarh', 'Nainital']),
    'Sen': ('West Bengal', ['Kolkata', 'Asansol', 'Siliguri', 'Durgapur', 'Bardhaman', 'Malda', 'Kharagpur', 'Haldia', 'Raiganj', 'Baharampur']),
    'Chakraborty': ('West Bengal', ['Kolkata', 'Asansol', 'Siliguri', 'Durgapur', 'Bardhaman', 'Malda', 'Kharagpur', 'Haldia', 'Raiganj', 'Baharampur']),
    'Yadav': ('Haryana', ['Faridabad', 'Gurgaon', 'Panipat', 'Ambala', 'Yamunanagar', 'Rohtak', 'Hisar', 'Karnal', 'Sonipat', 'Panchkula']),
    'Sinha': ('Bihar', ['Patna', 'Gaya', 'Bhagalpur', 'Muzaffarpur', 'Purnia', 'Darbhanga', 'Bihar Sharif', 'Arrah', 'Begusarai', 'Katihar']),
    'Bose': ('West Bengal', ['Kolkata', 'Asansol', 'Siliguri', 'Durgapur', 'Bardhaman', 'Malda', 'Kharagpur', 'Haldia', 'Raiganj', 'Baharampur']),
    'Chatterjee': ('West Bengal', ['Kolkata', 'Asansol', 'Siliguri', 'Durgapur', 'Bardhaman', 'Malda', 'Kharagpur', 'Haldia', 'Raiganj', 'Baharampur']),
    'Malhotra': ('Delhi', ['Delhi']),
    'Ghosh': ('West Bengal', ['Kolkata', 'Asansol', 'Siliguri', 'Durgapur', 'Bardhaman', 'Malda', 'Kharagpur', 'Haldia', 'Raiganj', 'Baharampur']),
    'Pandey': ('Uttarakhand', ['Dehradun', 'Haridwar', 'Roorkee', 'Haldwani', 'Rudrapur', 'Rishikesh', 'Kashipur', 'Kotdwar', 'Pithoragarh', 'Nainital']),
    'Rana': ('Himachal Pradesh', ['Shimla', 'Mandi', 'Solan', 'Dharamsala', 'Baddi', 'Una', 'Nahan', 'Palampur', 'Kullu', 'Chamba']),
    'Naik': ('Goa', ['Panaji', 'Margao', 'Vasco da Gama', 'Mapusa', 'Ponda', 'Bicholim', 'Curchorem', 'Canacona', 'Sanquelim', 'Valpoi']),
    'Murthy': ('Karnataka', ['Bangalore', 'Mysore', 'Mangalore', 'Hubli', 'Dharwad', 'Belgaum', 'Shimoga', 'Tumkur', 'Bijapur', 'Gulbarga']),
    'Krishnan': ('Karnataka', ['Bangalore', 'Mysore', 'Mangalore', 'Hubli', 'Dharwad', 'Belgaum', 'Shimoga', 'Tumkur', 'Bijapur', 'Gulbarga']),
    'Swamy': ('Karnataka', ['Bangalore', 'Mysore', 'Mangalore', 'Hubli', 'Dharwad', 'Belgaum', 'Shimoga', 'Tumkur', 'Bijapur', 'Gulbarga']),
    'Selvam': ('Karnataka', ['Bangalore', 'Mysore', 'Mangalore', 'Hubli', 'Dharwad', 'Belgaum', 'Shimoga', 'Tumkur', 'Bijapur', 'Gulbarga']),
    'Naidu': ('Karnataka', ['Bangalore', 'Mysore', 'Mangalore', 'Hubli', 'Dharwad', 'Belgaum', 'Shimoga', 'Tumkur', 'Bijapur', 'Gulbarga']),
    "Chandrakar": ("Odisha", ["Bhubaneswar", "Cuttack", "Rourkela", "Puri", "Sambalpur"]),
    "Panigrahi": ("Odisha", ["Bhubaneswar", "Cuttack", "Rourkela", "Puri", "Sambalpur"]),
    "Pradhan": ("Odisha", ["Bhubaneswar", "Cuttack", "Rourkela", "Puri", "Sambalpur"]),
    "Dey": ("Odisha", ["Bhubaneswar", "Cuttack", "Rourkela", "Puri", "Sambalpur"]),
    "Dash": ("Odisha", ["Bhubaneswar", "Cuttack", "Rourkela", "Puri", "Sambalpur"]),

    # Maharashtra
    "Gore": ("Maharashtra", ["Mumbai", "Pune", "Nagpur", "Nashik", "Aurangabad"]),
    "Kale": ("Maharashtra", ["Mumbai", "Pune", "Nagpur", "Nashik", "Aurangabad"]),
    "Sanvale": ("Maharashtra", ["Mumbai", "Pune", "Nagpur", "Nashik", "Aurangabad"]),
    "Sonavane": ("Maharashtra", ["Mumbai", "Pune", "Nagpur", "Nashik", "Aurangabad"]),
    "Chakole": ("Maharashtra", ["Mumbai", "Pune", "Nagpur", "Nashik", "Aurangabad"]),
    "Talpade": ("Maharashtra", ["Mumbai", "Pune", "Nagpur", "Nashik", "Aurangabad"]),
    "Shikhre": ("Maharashtra", ["Mumbai", "Pune", "Nagpur", "Nashik", "Aurangabad"]),
    "Gokhale": ("Maharashtra", ["Mumbai", "Pune", "Nagpur", "Nashik", "Aurangabad"]),
    "Chavan": ("Maharashtra", ["Mumbai", "Pune", "Nagpur", "Nashik", "Aurangabad"]),
    "Tiwari": ('Bihar', ['Patna', 'Gaya', 'Bhagalpur', 'Muzaffarpur', 'Purnia', 'Darbhanga', 'Bihar Sharif', 'Arrah', 'Begusarai', 'Katihar']),
    "Bajpai": ("Uttar Pradesh", ["Lucknow", "Kanpur", "Varanasi", "Agra"]),
    "Singh": ("Uttar Pradesh", ["Lucknow", "Kanpur", "Varanasi", "Agra"]),
    "Singhania": ("Uttar Pradesh", ["Lucknow", "Kanpur", "Varanasi", "Agra"]),
    "Galgotia": ("Uttar Pradesh", ["Lucknow", "Kanpur", "Varanasi", "Agra"]),
    "Rajpoot": ("Uttar Pradesh", ["Lucknow", "Kanpur", "Varanasi", "Agra"]),

    # Rajasthan
    "Thapar": ("Rajasthan", ["Jaipur", "Jodhpur", "Udaipur", "Kota", "Ajmer"]),
    "Pahariya": ("Rajasthan", ["Jaipur", "Jodhpur", "Udaipur", "Kota", "Ajmer"]),
    "Chaudhary": ("Rajasthan", ["Jaipur", "Jodhpur", "Udaipur", "Kota", "Ajmer"]),
    "Bhagat": ("Rajasthan", ["Jaipur", "Jodhpur", "Udaipur", "Kota", "Ajmer"]),
    "Pawar": ("Rajasthan", ["Jaipur", "Jodhpur", "Udaipur", "Kota", "Ajmer"]),

    # Haryana
    "Rajput": ("Haryana", ["Gurgaon", "Faridabad", "Panipat", "Hisar", "Ambala"]),
    "Rastogi": ("Haryana", ["Gurgaon", "Faridabad", "Panipat", "Hisar", "Ambala"]),

    # Karnataka
    "Shetty": ("Karnataka", ["Bangalore", "Mangalore", "Udupi", "Hubli", "Mysore"]),

    # Madhya Pradesh
    "Kulkarni": ("Madhya Pradesh", ["Indore", "Bhopal", "Gwalior", "Jabalpur", "Ujjain"]),

    # Uttarakhand
    "Kashyap": ("Uttarakhand", ["Dehradun", "Haridwar", "Nainital", "Rishikesh", "Almora"]),

    # West Bengal
    "Dwivedi": ("West Bengal", ["Kolkata", "Howrah", "Durgapur", "Siliguri", "Asansol"]),
    "Pathak": ('Bihar', ['Patna', 'Gaya', 'Bhagalpur', 'Muzaffarpur', 'Purnia', 'Darbhanga', 'Bihar Sharif', 'Arrah', 'Begusarai', 'Katihar']),
    "Garg": ("West Bengal", ["Kolkata", "Howrah", "Durgapur", "Siliguri", "Asansol"]),
    "Lal": ("West Bengal", ["Kolkata", "Howrah", "Durgapur", "Siliguri", "Asansol"]),
    "Thakur": ('Bihar', ['Patna', 'Gaya', 'Bhagalpur', 'Muzaffarpur', 'Purnia', 'Darbhanga', 'Bihar Sharif', 'Arrah', 'Begusarai', 'Katihar']),
    "Roy": ("West Bengal", ["Kolkata", "Howrah", "Durgapur", "Siliguri", "Asansol"]),

    # West Bengal (Continued)
    "Dasgupta": ("West Bengal", ["Kolkata", "Howrah", "Durgapur", "Siliguri", "Asansol"]),
    "Mukherjee": ("West Bengal", ["Kolkata", "Howrah", "Durgapur", "Siliguri", "Asansol"]),
    "Majumdar": ("West Bengal", ["Kolkata", "Howrah", "Durgapur", "Siliguri", "Asansol"]),
    "Kar": ("West Bengal", ["Kolkata", "Howrah", "Durgapur", "Siliguri", "Asansol"]),
    "Sarkar": ("West Bengal", ["Kolkata", "Howrah", "Durgapur", "Siliguri", "Asansol"]),

    # North Indian states (Additional names)
    "Chaudhari": ("Haryana", ["Gurgaon", "Faridabad", "Panipat", "Hisar", "Ambala"]),
    "Salvi": ("Rajasthan", ["Jaipur", "Jodhpur", "Udaipur", "Kota", "Ajmer"]),
    "Gadre": ("Maharashtra", ["Mumbai", "Pune", "Nagpur", "Nashik", "Aurangabad"]),
    "Khatri": ("Punjab", ["Amritsar", "Ludhiana", "Jalandhar", "Patiala", "Bathinda"]),
    "Mohanty": ("Odisha", ["Bhubaneswar", "Cuttack", "Rourkela", "Puri", "Sambalpur"]),
    "Saxena": ("Uttar Pradesh", ["Lucknow", "Kanpur", "Varanasi",  "Agra"]),
    "Shroff": ("Gujarat", ["Ahmedabad", "Surat", "Vadodara", "Rajkot", "Bhavnagar"]),
    "Deshmukh": ("Maharashtra", ["Mumbai", "Pune", "Nagpur", "Nashik", "Aurangabad"]),
    "Gaikwad": ("Maharashtra", ["Mumbai", "Pune", "Nagpur", "Nashik", "Aurangabad"]),
    "Raj": ("Rajasthan", ["Jaipur", "Jodhpur", "Udaipur", "Kota", "Ajmer"]),
    "Acharya": ("Gujarat", ["Ahmedabad", "Surat", "Vadodara", "Rajkot", "Bhavnagar"]),
    "Kohli": ("Punjab", ["Amritsar", "Ludhiana", "Jalandhar", "Patiala", "Bathinda"]),
    "Sethi": ("Punjab", ["Amritsar", "Ludhiana", "Jalandhar", "Patiala", "Bathinda"]),
    "Puri": ("Punjab", ["Amritsar", "Ludhiana", "Jalandhar", "Patiala", "Bathinda"]),
    "Mahajan": ("Punjab", ["Amritsar", "Ludhiana", "Jalandhar", "Patiala", "Bathinda"]),
    "Bhonsle": ("Maharashtra", ["Mumbai", "Pune", "Nagpur", "Nashik", "Aurangabad"]),
    "Bhadauria": ("Uttar Pradesh", ["Lucknow", "Kanpur", "Varanasi",  "Agra"]),
    "Ghoshal": ("West Bengal", ["Kolkata", "Howrah", "Durgapur", "Siliguri", "Asansol"]),
    "Kadam": ("Maharashtra", ["Mumbai", "Pune", "Nagpur", "Nashik", "Aurangabad"]),
    "Kaur": ("Punjab", ["Amritsar", "Ludhiana", "Jalandhar", "Patiala", "Bathinda"]),
    "Pal": ("Uttar Pradesh", ["Lucknow", "Kanpur", "Varanasi",  "Agra"]),
    "Rawat": ("Uttarakhand", ["Dehradun", "Haridwar", "Nainital", "Rishikesh", "Almora"]),
    "Kapoor": ("Punjab", ["Amritsar", "Ludhiana", "Jalandhar", "Patiala", "Bathinda"]),
    "Grover": ("Punjab", ["Amritsar", "Ludhiana", "Jalandhar", "Patiala", "Bathinda"]),
    "Kalra": ("Punjab", ["Amritsar", "Ludhiana", "Jalandhar", "Patiala", "Bathinda"]),
    "Khurana": ("Punjab", ["Amritsar", "Ludhiana", "Jalandhar", "Patiala", "Bathinda"]),
    "Vohra": ("Punjab", ["Amritsar", "Ludhiana", "Jalandhar", "Patiala", "Bathinda"]),
    "Sidhu": ("Punjab", ["Amritsar", "Ludhiana", "Jalandhar", "Patiala", "Bathinda"]),

}

# Function to generate a random mobile number
def generate_mobile_number():
    return f"{random.choice([6,7, 8, 9])}{random.randint(1000000000, 9999999999)}"[0:10]


# Function to get state and city based on surname
def get_state_city(surname):
    state, cities = surnames_to_states.get(surname, ("Unknown", ["Unknown City"]))
    return state, random.choice(cities)

# Generate 3900 new unique names avoiding duplicates
new_entries = []
count=0
csf=1
csmf=1
cmf=1
clf=1
cxf=1
csm=1
csmm=1
cmm=1
clm=1
cxm=1
for i in range(1):  
    for j in range(1, 124):
        if 93 <= j <= 117: 
            while True:
                name = f"{random.choice(new_female_names)} {random.choice(surnames)}"
                plan = f"{random.choice(Plan)}"
                food = f"{random.choice(Food)}"
                mode = f"{random.choice(Mode)}"
                group = f"{random.choice(Group)}"
                travel = f"{random.choice(Travel)}"
                religious = f"{random.choice(ReligiousItems)}"
                recreation = f"{random.choice(Recreation)}"
                shopping = f"{random.choice(Shopping)}"
                others = f"{random.choice(Others)}"
                fav = f"{random.choice(Visits)}"
                next = f"{random.choice(Visits)}"
                last = f"{random.choice(Visits)}"
                if next != last:
                    if name not in previous_names:
                        previous_names.add(name)
                        surname = name.split()[-1]
                        state, city = get_state_city(surname)
                        if 95<=j<=99:
                            new_entries.append([name, "19-25", "Female", generate_mobile_number(), f"{img_path}smf{csmf}.jpg", state, city,next, last, fav, plan, mode, group, travel, food, religious, recreation, shopping, others])
                            csmf+=1
                        elif 100<=j<=103:
                            new_entries.append([name, "26-40", "Female", generate_mobile_number(), f"{img_path}mf{cmf}.jpg", state, city,next, last, fav, plan, mode, group, travel, food, religious, recreation, shopping, others])
                            cmf+=1
                        elif 93<=j<=94:
                            new_entries.append([name, "Under 18", "Female", generate_mobile_number(), f"{img_path}sf{csf}.jpg", state, city,next, last, fav, plan, mode, group, travel, food, religious, recreation, shopping, others])
                            csf+=1
                        elif 104<=j<=117:
                            new_entries.append([name, "41-60", "Female", generate_mobile_number(), f"{img_path}lf{clf}.jpg", state, city,next, last, fav, plan, mode, group, travel, food, religious, recreation, shopping, others])
                            clf+=1
                        # elif 118<=j<=123:
                        #     new_entries.append([name, "Above 60", "Female", generate_mobile_number()])
                        count= count+1
                        print(count)
                        break
        elif  1<= j <= 82 :  # Female entries
            while True:
                name = f"{random.choice(new_male_names)} {random.choice(surnames)}"
                plan = f"{random.choice(Plan)}"
                food = f"{random.choice(Food)}"
                mode = f"{random.choice(Mode)}"
                group = f"{random.choice(Group)}"
                travel = f"{random.choice(Travel)}"
                religious = f"{random.choice(ReligiousItems)}"
                recreation = f"{random.choice(Recreation)}"
                shopping = f"{random.choice(Shopping)}"
                others = f"{random.choice(Others)}"
                fav = f"{random.choice(Visits)}"
                next = f"{random.choice(Visits)}"
                last = f"{random.choice(Visits)}"
                if next != last:
                    if name not in previous_names:
                        previous_names.add(name)
                        surname = name.split()[-1]
                        state, city = get_state_city(surname)
                        if 15<=j<=23:
                            new_entries.append([name, "19-25", "Male", generate_mobile_number(), f"{img_path}smm{csmm}.jpg", state, city,next, last, fav, plan, mode, group, travel, food, religious, recreation, shopping, others])
                            csmm+=1
                        elif 24<=j<=53:
                            new_entries.append([name, "26-40", "Male", generate_mobile_number(), f"{img_path}smm{cmm}.jpg", state, city,next, last, fav, plan, mode, group, travel, food, religious, recreation, shopping, others])
                            cmm+=1
                        elif 1<=j<=14:
                            new_entries.append([name, "Under 18", "Male", generate_mobile_number(), f"{img_path}sm{csm}.jpg", state, city, next, last, fav, plan, mode, group, travel, food, religious, recreation, shopping, others])
                            csm+=1
                        elif 54<=j<=82:
                            new_entries.append([name, "41-60", "Male", generate_mobile_number(), f"{img_path}lm{clm}.jpg", state, city, next, last, fav, plan, mode, group, travel, food, religious, recreation, shopping, others])
                            clm+=1
                        # elif 83<=j<=92:
                        #     new_entries.append([name, "Above 60", "Female", generate_mobile_number()])
                        count = count+1
                        print(count)
                        break
        elif 83<=j<=92:
            while True:
                name = f"{random.choice(old_male_names)} {random.choice(surnames)}"
                plan = f"{random.choice(Plan)}"
                food = f"{random.choice(Food)}"
                mode = f"{random.choice(Mode)}"
                group = f"{random.choice(Group)}"
                travel = f"{random.choice(Travel)}"
                religious = f"{random.choice(ReligiousItems)}"
                recreation = f"{random.choice(Recreation)}"
                shopping = f"{random.choice(Shopping)}"
                others = f"{random.choice(Others)}"
                fav = f"{random.choice(Visits)}"
                next = f"{random.choice(Visits)}"
                last = f"{random.choice(Visits)}"
                if next != last:
                    if name not in previous_names:
                        previous_names.add(name)
                        surname = name.split()[-1]
                        state, city = get_state_city(surname)
                        new_entries.append([name, "Above 60", "Male", generate_mobile_number(), f"{img_path}xm{cxm}.jpg", state, city,next, last, fav, plan, mode, group, travel, food, religious, recreation, shopping, others])
                        cxm+=1
                        count = count+1
                        print(count)
                        break
        elif 118<=j<=123:
            while True:
                name = f"{random.choice(old_female_names)} {random.choice(surnames)}"
                plan = f"{random.choice(Plan)}"
                mode = f"{random.choice(Mode)}"
                food = f"{random.choice(Food)}"
                group = f"{random.choice(Group)}"
                travel = f"{random.choice(Travel)}"
                religious = f"{random.choice(ReligiousItems)}"
                recreation = f"{random.choice(Recreation)}"
                shopping = f"{random.choice(Shopping)}"
                others = f"{random.choice(Others)}"
                fav = f"{random.choice(Visits)}"
                next = f"{random.choice(Visits)}"
                last = f"{random.choice(Visits)}"
                if next != last:
                    if name not in previous_names:
                        previous_names.add(name)
                        surname = name.split()[-1]
                        state, city = get_state_city(surname)
                        new_entries.append([name, "Above 60", "Female", generate_mobile_number(), f"{img_path}xf{cxf}.jpg", state, city, next, last, fav, plan, mode, group, travel, food, religious, recreation, shopping, others])
                        cxf+=1
                        count = count+1
                        print(count)
                        break


# Write the new entries to a CSV file
# new_csv_file = "survey.xlsx"
# with open(new_csv_file, mode="w", newline="") as file:
#     writer = xlsx.writer(file)
#     writer.writerow(["Name", "Age", "Gender", "Mobile Number", "Image", "State", "City", "Plan", "Mode", "Group", "Travel", "Food", "ReligiousItems", "Recreation", "Shopping", "Others"])
#     writer.writerows(new_entries)

# new_csv_file
new_xlsx_file = "output.xlsx"  # Specify the name of the output file
workbook = openpyxl.Workbook()
sheet = workbook.active

header = ["Name", "Age", "Gender", "Mobile Number", "Image", "State", "City", "NextVisit", "LastVisit", "FavVisit", "Plan", "Mode", "Group", "Travel", "Food", "ReligiousItems", "Recreation", "Shopping", "Others"]
sheet.append(header)

for entry in new_entries:
    sheet.append(entry)

# Save the workbook to the specified file
workbook.save(new_xlsx_file)
