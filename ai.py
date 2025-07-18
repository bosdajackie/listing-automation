import pandas as pd
import json
import os
from openai import OpenAI
from dotenv import load_dotenv
import base64
load_dotenv()

def format_vehicle_lines_from_df(df, max_len=65):
    """Format vehicle compatibility data into readable lines"""
    formatted_lines = []

    for _, row in df.iterrows():
        make = str(row['Make']).upper()
        model = str(row['Model']).upper()
        year_range = str(row['Year']).strip()
        position = str(row['Position']).title().strip()
        extra = str(row['Engine']).strip()

        # Expand 2017-2019 to ['2017', '2018', '2019']
        if "-" in year_range:
            start, end = map(int, year_range.split("-"))
            years = [str(y) for y in range(start, end + 1)]
        else:
            years = [year_range.strip()]

        base = f"{make} {model} {position}"
        if extra and extra.upper() not in base.upper():
            base += f" {extra}"

        current_years = []
        for year in years:
            test_line = f"{' '.join(current_years + [year])} {base}"
            if len(test_line) <= max_len:
                current_years.append(year)
            else:
                if current_years:
                    formatted_lines.append(f"{' '.join(current_years)} {base}.")
                current_years = [year]

        if current_years:
            formatted_lines.append(f"{' '.join(current_years)} {base}.")
    
    return "\n".join(formatted_lines)

def ai_generate_title(category, selected_vehicle, vehicle_list):
    if "OPENAI_API_KEY" in os.environ:
        try:
            client = OpenAI()
            
            rules = (
                """
                    You are writing an 80-character product listing title for an online automotive parts marketplace.

                    Your goal is to maximize click-through rate and buyer interest by highlighting:
                    - The most important keywords (part type, vehicle name, model years)
                    - Popular search terms used by customers
                    - Key selling features (like OEM fit, ABS sensor, or bolt pattern if known)

                    Requirements:
                    - Do NOT exceed 80 characters
                    - Mention the vehicle make/model/year range only once (brief and effective)
                    - Include position (Front, Rear, Left, Right) if provided
                    - Include category (e.g., Wheel Bearing & Hub Assembly, Wheel Hub, Knuckle Assembly, Bearing, etc.)
                    - Use title case
                    - Avoid fluff or punctuation (no !)
                    - Never guess any technical details—only use what's fact checked
                """
            )

            prompt = (
                f"You are generating a product listing title for a {category}.\n"
                f"Here is the selected vehicle (if any): {selected_vehicle}\n"
                "If no selected vehicle was provided, please choose the most popular vehicle from the vehicle list according to sales volume in the given year range.\n"
                f"Here is the vehicle data:\n{vehicle_list}\n"
                f"Use the formatting rules I gave you earlier."
            )

            response = client.responses.create(
                model="gpt-4o",
                input=prompt,
                instructions=rules,
                temperature=0,
            )
            
            desc = response.output_text
            return desc
            
        except Exception as e:
            print(f"OpenAI API failed: {e}")
            return False
    else:
        return False

def ai_generate_short_description(df, category):
    """Generate short description from compatibility data"""
    # Build manual fallback description
    formatted = format_vehicle_lines_from_df(df)
    
    # Prepare data for AI
    rows = []
    for _, r in df.iterrows():
        rows.append(f"{r['Make']} {r['Model']} {r['Year']} {r['Position']} {r['Engine']}")
    vehicle_list = json.dumps(rows)

    # Try to use OpenAI API
    if "OPENAI_API_KEY" in os.environ:
        try:
            client = OpenAI()
            
            rules = (
                "You are generating a short e-commerce description for a wheel hub and bearing assembly based on vehicle compatibility data.\n\n"
                "Follow these exact instructions:\n"
                "- Each vehicle line must begin with space-separated years (e.g., 2014 2015 2016).\n"
                "- Then list Make, Model, and side or position (e.g., Front Right).\n"
                "- If the combined years + Make/Model/Position exceeds 65 characters, split the years across multiple lines. Each line must still repeat the full Make, Model, and Position.\n"
                "- No hyphens or en dashes in year ranges. List years explicitly.\n"
                "- DO NOT wrap mid-model or after make/model/position. Only wrap years to a new line.\n"
                "- Each output line must be **a complete line** followed by a newline character.\n"
                "- Make sure there is nothing before or after all the output (ex. no ```).\n"
                "- Never put two vehicle fitments on the same line.\n"
                "- Do NOT guess drive type, lug count, or features. Only include if confirmed and fact checked!!!\n"
                "- Make sure that each sentence that you generate is a new line.\n"
                "- After listing all vehicles, include exactly 4 final lines:\n"
                "  1. Line describing which sides/positions it fits (e.g., 'Front Left Right...').\n"
                "  2. Line describing what the part includes (e.g., 'Includes hub, bearing...').\n"
                "  3. Line summarizing verified physical or trim details (e.g., 'Fits AWD 5 Lug 5 Bolt 5 Stud.').\n"
                "  4. Final line: Compatibility summary, e.g., 'Compatible with 2 and 4 Door Compact Luxury Models'. This is the only line allowed to use the word 'Compatible'. Only include if confirmed and fact checked!!!\n"
            
                "Example Output:\n"
                "2014 2015 2016 BMW 228i RWD Front.\n"
                "2012 2013 2014 2015 2016 BMW 320i RWD Front.\n"
                "2014 2015 2016 BMW 328d RWD Front.\n"
                "2013 2014 2015 2016 BMW 328i RWD Front.\n"
                "2014 2015 BMW 335i RWD Front.\n"
                "2016 BMW 340i Base RWD Front.\n"
                "2014 2015 2016 BMW 428i RWD Front.\n"
                "2014 2015 2016 BMW 435i RWD Front.\n"
                "2013 2014 2015 BMW ActiveHybrid 3 Front.\n"
                "2014 2015 2016 BMW M235i RWD Front.\n"
                "Front Left Right, Front Left Side Right Side.\n"
                "Fits Base, Luxury, Sport, M Sport, and M trims.\n"
                "Wheel hub assembly with 12mm bolt mounting dimension.\n"
                "Fits RWD 5 Lug 5 Bolt 5 Stud.\n"
                "Compatible with 2 and 4 Door Compact Luxury Models.\n"
                "BMW 328I 2013-2016 Front RWD, 12mm Bolt Mounting Dimension\n"
                "BMW 340I 2016 Front 340i Base Model, 12mm Bolt Mounting Dimension\n"
            )

            prompt = (
                f"You are generating a short e-commerce description for a {category}.\n"
                f"Here is the vehicle data:\n{vehicle_list}\n"
                f"Use the formatting rules I gave you earlier."
            )

            response = client.responses.create(
                model="gpt-4o",
                input=prompt,
                instructions=rules,
                temperature=0,
            )
            
            desc = response.output_text
            return desc
            
        except Exception as e:
            print(f"OpenAI API failed: {e}")
            return formatted
    else:
        return formatted

def ai_generate_long_description(df, part_number, alternate_numbers):
    """Generate long description from compatibility data"""
    # Prepare data for AI
    rows = []
    for _, r in df.iterrows():
        rows.append(f"{r['Make']} {r['Model']} {r['Year']} {r['Position']} {r['Engine']}")
    vehicle_list = json.dumps(rows)

    # Try to use OpenAI API
    if "OPENAI_API_KEY" in os.environ:
        try:
            client = OpenAI()
            
            rules = (
                "You are generating a long-format, keyword-rich e-commerce description for an automotive part "
                "(usually a wheel bearing, hub assembly, or knuckle assembly).\n\n"
                "You are given:\n"
                "- A list of compatible vehicles (with year, make, model, position, and trim/engine details)\n"
                "- A list of alternate part numbers (OE, OEM, aftermarket, and supplier SKUs)\n\n"
                "Your job is to generate a single long line of keywords, without line breaks, for use in automotive marketplaces.\n\n"
                "Output Format:\n"
                "- Begin with all part numbers (main part + alternates)\n"
                "- Follow with compatible years (each year individually, e.g., 2012 2013 2014)\n"
                "- Then add all makes and models (e.g., BMW 328i)\n"
                "- Then add all relevant positions (Front, Rear, Left, Right, Driver, Passenger)\n"
                "- Include drive types or trims if explicitly provided (e.g., RWD, AWD, Base, M Sport)\n"
                "- Include any mounting specs (e.g., 12mm Bolt Mounting Dimension)\n"
                "- Include marketing keywords and synonyms (e.g., Wheel Hub, Hub Assembly, OE Replacement, Repair Kit, etc.)\n"
                "- End with relevant features or selling points (e.g., Pre-Greased, With ABS-- only if fact-checked, Corrosion-Resistant, Precision-Machined, 1-year warranty, 30-day returns)\n"
                "- Make it as detailed as possible since the point of long desc is to get more people to see this listing\n\n"
                "Constraints:\n"
                "- DO NOT GUESS any details make sure that they are all fact checked\n"
                "- Use title case for vehicle models and trims, but part numbers stay UPPERCASE\n"
                "- Use spaces to separate words, no commas or periods\n"
                "- No repeated values, combine overlapping data naturally (e.g., 2012-2016 → 2012 2013 2014 2015 2016)\n\n"
                "Output must be a single continuous string of keywords, separated only by spaces.\n\n"
                
                "Example input for wheel bearing & hub assembly:\n"
                "513359, 31206794850, 31206857230, 31206867256\n"
                "BMW	228I	2014-2016	Front	RWD, 12mm Bolt Mounting Dimension\n"
                "BMW	320I	2012-2016	Front	RWD, 12mm Bolt Mounting Dimension\n"
                "BMW	328D	2014-2016	Front	RWD, 12mm Bolt Mounting Dimension\n"
                "BMW	328I	2013-2016	Front	RWD, 12mm Bolt Mounting Dimension\n"
                "BMW	335I	2014-2015	Front	RWD, 12mm Bolt Mounting Dimension\n"
                "BMW	340I	2016	Front	340i Base Model, 12mm Bolt Mounting Dimension\n"
                "BMW	428I	2014-2016	Front	RWD, 12mm Bolt Mounting Dimension\n"
                "BMW	435I	2014-2016	Front	RWD, 12mm Bolt Mounting Dimension\n"
                "BMW	ACTIVEHYBRID 3	2013-2015	Front	12mm Bolt Mounting Dimension\n"
                "BMW	M235I	2014-2016	Front	RWD, 12mm Bolt Mounting Dimension\n\n"

                "Example output:\n"
                "513359 31206794850 31206857230 31206867256 WA513359 TRQ BHA513359 HUB513359 WH513359 H513359 SP513359 "
                "2012 2013 2014 2015 2016 12 13 14 15 16 BMW 228i 320i 328d 328i 335i 340i 428i 435i ActiveHybrid 3 M235i "
                "Front Wheel Drive RWD Rear-Wheel Drive Only 12mm Bolt Mounting Dimension Front Left Right Driver Passenger "
                "L4 2.0L L6 3.0L Turbocharged Sedan Coupe Convertible Base Luxury Sport M Sport xDrive Where Noted "
                "5 Bolt 5 Stud 5 Lug 4 Flange With ABS With Tone Ring Without Sensor Wire Wheel Bearing and Hub Assembly "
                "OE Replacement HD Heavy Duty 1 Year Warranty OEM Specification OE Replacement OE Performance "
                "Wheel Hub Hub Assembly Wheel Bearing Hub Bearing Hub Unit Axle Bearing Bearing Hub Wheel Hub Assembly Auto Hub Assembly Car Hub Assembly "
                "Grease-packed Bearings Corrosion-Resistant Pre-Greased Precision-Machined Ready to Install Front Axle Front Hub Assembly Front Axle Bearing Assembly"
            )

            prompt = (
                f"You are generating a long e-commerce description for part number {part_number}\n"
                f"Here are the alternate numbers for that part: {alternate_numbers}\n"
                f"Here is the vehicle data:\n{vehicle_list}\n"
                f"Use the formatting rules I gave you earlier."
            )

            response = client.responses.create(
                model="gpt-4o",
                input=prompt,
                instructions=rules,
                temperature=0,
            )
            
            desc = response.output_text
            return desc
            
        except Exception as e:
            print(f"OpenAI API failed: {e}")
            # Return a basic fallback description
            return f"Part Number {part_number} - Compatible with multiple vehicle models. Please see compatibility chart for details."
    else:
        return f"Part Number {part_number} - Compatible with multiple vehicle models. Please see compatibility chart for details."
    
def ai_generate_image(selected_vehicle, vehicle_list):
    # Try to use OpenAI API
    if "OPENAI_API_KEY" in os.environ:
        try:
            client = OpenAI()

            prompt = (
                "You are generating a high-quality vehicle image.\n"
                "Constraints:\n"
                "- Show the vehicle from a front 3/4 view (front quarter angle), slightly elevated.\n"
                "- The vehicle must be centered, realistic, high-resolution, and uncropped, with white background\n"
                "- Show the full left and right sides, front bumper, and front wheels clearly. The vehicle should NOT be cropped at all\n"
                "- Match the factory trim and shape. Remove any license plates.\n"
                "- Use a professional, clean, studio-lit style.\n"
                "- Use factory-correct proportions and avoid cartoon or artistic rendering.\n"
                "- Preferred colors: red, white, silver, black, and blue (rotate if repeating).\n\n"
                f"Here is the selected vehicle (if any): {selected_vehicle}\n"
                "If no selected vehicle was provided, please choose the most popular vehicle from the vehicle list according to sales volume in the given year range.\n"
                f"Here is the vehicle data:\n{vehicle_list}\n"
            )

            result = client.images.generate(
                model="gpt-image-1",
                prompt=prompt
            )

            image_base64 = result.data[0].b64_json
            image_bytes = base64.b64decode(image_base64)
            currentPath = os.getcwd()
            save_path = os.path.join(os.path.join(currentPath, "results"), "vehicle_image.jpg")

            os.makedirs(os.path.dirname(save_path), exist_ok=True)
            with open(save_path, "wb") as f:
                f.write(image_bytes)
            print(f"Image saved to {save_path}")
            return save_path
        except Exception as e:
            print(f"Image generation failed: {e}")
            return False