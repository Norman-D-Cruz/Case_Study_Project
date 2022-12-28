# Will be used for logging in
linkedin_username = 'xxxxxxxx'            # LinkedIn username
linkedin_password = 'xxxxxxxx'            # Your LinkedIn password

country_list = ['Afghanistan', 'Aland Islands', 'Albania', 'Algeria', 'American Samoa', 'Andorra', 'Angola', 'Anguilla', 'Antarctica', 'Antigua and Barbuda', 'Argentina', 'Armenia', 'Aruba', 'Australia', 'Austria', 'Azerbaijan', 'Bahamas', 'Bahrain', 'Bangladesh', 'Barbados', 'Belarus', 'Belgium', 'Belize', 'Benin', 'Bermuda', 'Bhutan', 'Bolivia, Plurinational State of', 'Bonaire, Sint Eustatius and Saba', 'Bosnia and Herzegovina', 'Botswana', 'Bouvet Island', 'Brazil', 'British Indian Ocean Territory', 'Brunei Darussalam', 'Bulgaria', 'Burkina Faso', 'Burundi', 'Cambodia', 'Cameroon', 'Canada', 'Cape Verde', 'Cayman Islands', 'Central African Republic', 'Chad', 'Chile', 'China', 'Christmas Island', 'Cocos Islands', 'Colombia', 'Comoros', 'Congo', 'Congo, The Democratic Republic of the', 'Cook Islands', 'Costa Rica', "Côte d'Ivoire", 'Croatia', 'Cuba', 'Curaçao', 'Cyprus', 'Czech Republic', 'Denmark', 'Djibouti', 'Dominica', 'Dominican Republic', 'Ecuador', 'Egypt', 'El Salvador', 'Equatorial Guinea', 'Eritrea', 'Estonia', 'Ethiopia', 'Falkland Islands', 'Faroe Islands', 'Fiji', 'Finland', 'France', 'French Guiana', 'French Polynesia', 'French Southern Territories', 'Gabon', 'Gambia', 'Georgia', 'Germany', 'Ghana', 'Gibraltar', 'Greece', 'Greenland', 'Grenada', 'Guadeloupe', 'Guam', 'Guatemala', 'Guernsey', 'Guinea', 'Guinea-Bissau', 'Guyana', 'Haiti', 'Heard Island and McDonald Islands', 'Holy See', 'Honduras', 'Hong Kong', 'Hungary', 'Iceland', 'India', 'Indonesia', 'Iran, Islamic Republic of', 'Iraq', 'Ireland', 'Isle of Man', 'Israel', 'Italy', 'Jamaica', 'Japan', 'Jersey', 'Jordan', 'Kazakhstan', 'Kenya', 'Kiribati', "Korea, Democratic People's Republic of", 'Korea, Republic of', 'Kuwait', 'Kyrgyzstan', "Lao People's Democratic Republic", 'Latvia', 'Lebanon', 'Lesotho', 'Liberia', 'Libya', 'Liechtenstein', 'Lithuania', 'Luxembourg', 'Macao', 'Macedonia, Republic of', 'Madagascar', 'Malawi', 'Malaysia', 'Maldives', 'Mali', 'Malta', 'Marshall Islands', 'Martinique', 'Mauritania', 'Mauritius', 'Mayotte', 'Mexico', 'Micronesia, Federated States of', 'Moldova, Republic of', 'Monaco', 'Mongolia', 'Montenegro', 'Montserrat', 'Morocco', 'Mozambique', 'Myanmar', 'Namibia', 'Nauru', 'Nepal', 'Netherlands', 'New Caledonia', 'New Zealand', 'Nicaragua', 'Niger', 'Nigeria', 'Niue', 'Norfolk Island', 'Northern Mariana Islands', 'Norway', 'Oman', 'Pakistan', 'Palau', 'Palestinian Territory, Occupied', 'Panama', 'Papua New Guinea', 'Paraguay', 'Peru', 'Philippines', 'Pitcairn', 'Poland', 'Portugal', 'Puerto Rico', 'Qatar', 'Réunion', 'Romania', 'Russian Federation', 'Rwanda', 'Saint Barthélemy', 'Saint Helena, Ascension and Tristan da Cunha', 'Saint Kitts and Nevis', 'Saint Lucia', 'Saint Martin', 'Saint Pierre and Miquelon', 'Saint Vincent and the Grenadines', 'Samoa', 'San Marino', 'Sao Tome and Principe', 'Saudi Arabia', 'Senegal', 'Serbia', 'Seychelles', 'Sierra Leone', 'Singapore', 'Sint Maarten', 'Slovakia', 'Slovenia', 'Solomon Islands', 'Somalia', 'South Africa', 'South Georgia and the South Sandwich Islands', 'Spain', 'Sri Lanka', 'Sudan', 'Suriname', 'South Sudan', 'Svalbard and Jan Mayen', 'Swaziland', 'Sweden', 'Switzerland', 'Syrian Arab Republic', 'Taiwan, Province of China', 'Tajikistan', 'Tanzania, United Republic of', 'Thailand', 'Timor-Leste', 'Togo', 'Tokelau', 'Tonga', 'Trinidad and Tobago', 'Tunisia', 'Turkey', 'Turkmenistan', 'Turks and Caicos Islands', 'Tuvalu', 'Uganda', 'Ukraine', 'United Arab Emirates', 'United Kingdom', 'United States', 'United States Minor Outlying Islands', 'Uruguay', 'Uzbekistan', 'Vanuatu', 'Venezuela, Bolivarian Republic of', 'Viet Nam', 'Virgin Islands, British', 'Virgin Islands, U.S.', 'Wallis and Futuna', 'Yemen', 'Zambia', 'Zimbabwe']

capital_list = ["Abidjan", "Abu Dhabi", "Abuja", "Accra", "Adamstown", "Addis Ababa", "Aden", "Algiers", "Alofi", "Amman", "Amsterdam", "Andorra la Vella", "Ankara", "Antananarivo", "Apia", "Ashgabat", "Asmara", "Astana", "Asunción", "Athens", "Avarua", "Baghdad", "Baku", "Bamako", "Bandar Seri Begawan", "Bangkok", "Bangui", "Banjul", "Basseterre", "Beijing", "Beirut", "Belgrade", "Belmopan", "Berlin", "Bern", "Bishkek", "Bissau", "Bloemfontein", "Bogotá", "Brades", "Brasília", "Bratislava", "Brazzaville", "Bridgetown", "Brussels", "Bucharest", "Budapest", "Buenos Aires", "Bujumbura", "Cairo", "Camp Thunder Cove", "Canberra", "Cape Town", "Caracas", "Castries", "Cetinje", "Charlotte Amalie", "Chișinău", "Cockburn Town", "Colombo", "Conakry", "Copenhagen", "Cotonou", "Dakar", "Damascus", "Dar es Salaam", "Dhaka", "Dili", "Djibouti", "Dodoma", "Doha", "Donetsk", "Douglas", "Dublin", "Dushanbe", "Flying Fish Cove", "Freetown", "Funafuti", "Gaborone", "George Town", "Georgetown", "Georgetown", "Gibraltar", "Gitega", "Guatemala City", "Gustavia", "Hagåtña", "Hamilton", "Hanoi", "Harare", "Hargeisa", "Havana", "Helsinki", "Honiara", "Islamabad", "Jakarta", "Jamestown", "Jerusalem", "Jerusalem", "Juba", "Kabul", "Kampala", "Kathmandu", "Khartoum", "Kigali", "King Edward Point", "Kingston", "Kingston", "Kingstown", "Kinshasa", "Kuala Lumpur", "Kuwait City", "Kyiv", "La Paz", "Laâyoune", "Libreville", "Lilongwe", "Lima", "Lisbon", "Ljubljana", "Lobamba", "Lomé", "London", "Luanda", "Luhansk", "Lusaka", "Luxembourg", "Madrid", "Majuro", "Malabo", "Malé", "Managua", "Manama", "Manila", "Maputo", "Mariehamn", "Marigot", "Maseru", "Mata Utu", "Mbabane", "Mexico City", "Minsk", "Mogadishu", "Monaco", "Monrovia", "Montevideo", "Moroni", "Moscow", "Muscat", "N'Djamena", "Nairobi", "Nassau", "Naypyidaw", "New Delhi", "Ngerulmud", "Niamey", "Nicosia", "Nouakchott", "Nouméa", "Nuuk", "Oranjestad", "Oslo", "Ottawa", "Ouagadougou", "Pago Pago", "Palikir", "Panama City", "Papeete", "Paramaribo", "Paris", "Philipsburg", "Phnom Penh", "Plymouth", "Podgorica", "Port Louis", "Port Moresby", "Port of Spain", "Port Vila", "Port-au-Prince", "Porto-Novo", "Prague", "Praia", "Pretoria", "Pristina", "Putrajaya", "Pyongyang", "Quito", "Rabat", "Ramallah", "Reykjavík", "Riga", "Riyadh", "Road Town", "Rome", "Roseau", "Rothera", "Saipan", "San José", "San Juan", "San Marino", "San Salvador", "Sana'a", "Santiago", "Santo Domingo", "São Tomé", "Sarajevo", "Sejong City", "Seoul", "Singapore", "Skopje", "Sofia", "South Tarawa", "Sri Jayawardenepura Kotte", "St. George's", "St. Helier", "St. John's", "St. Peter Port", "St. Pierre", "Stanley", "Stepanakert", "Stockholm", "Sucre", "Sukhumi", "Suva", "Taipei", "Tallinn", "Tashkent", "Tbilisi", "Tegucigalpa", "Tehran", "The Hague", "The Valley", "Thimphu", "Tifariti", "Tirana", "Tiraspol", "Tokyo", "Tórshavn", "Tripoli", "Tskhinvali", "Tunis", "Ulaanbaatar", "Vaduz", "Valletta", "Valparaíso", "Vatican City", "Victoria", "Vienna", "Vientiane", "Vilnius", "Warsaw", "Washington", " D.C.", "Wellington", "West Island", "Willemstad", "Windhoek", "Yamoussoukro", "Yaoundé", "Yaren", "Yerevan", "Zagreb"]

columns_list = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ"]