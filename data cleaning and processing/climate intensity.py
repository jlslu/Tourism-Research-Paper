import pandas as pd
import numpy as np
import requests
import json
from datetime import datetime, timedelta
import time

class CITCalculator:
    """
    Climate Index for Tourism (CIT) Calculator
    Based on de Freitas, Scott, and McBoyle (2008) methodology
    CIT = 6.4 + 0.4*TSN - 0.281*TSN²
    """
    
    def __init__(self):
        # US State coordinates (capital cities as representatives)
        self.us_states = {
            'Alabama': (32.3617, -86.2792),
            'Alaska': (61.2181, -149.9003),
            'Arizona': (33.4484, -112.0740),
            'Arkansas': (34.7465, -92.2896),
            'California': (38.5767, -121.4934),
            'Colorado': (39.7392, -104.9903),
            'Connecticut': (41.5978, -72.7554),
            'Delaware': (39.1612, -75.5264),
            'Florida': (30.4518, -84.27277),
            'Georgia': (33.76, -84.39),
            'Hawaii': (21.30895, -157.826182),
            'Idaho': (43.2081, -116.2146),
            'Illinois': (39.78325, -89.650373),
            'Indiana': (39.790942, -86.147685),
            'Iowa': (41.590939, -93.620866),
            'Kansas': (39.04, -95.69),
            'Kentucky': (38.197274, -84.86311),
            'Louisiana': (30.45809, -91.140229),
            'Maine': (44.323535, -69.765261),
            'Maryland': (38.972945, -76.501157),
            'Massachusetts': (42.2352, -71.0275),
            'Michigan': (42.354558, -84.955255),
            'Minnesota': (44.95, -93.094),
            'Mississippi': (32.354668, -90.178217),
            'Missouri': (38.572954, -92.189283),
            'Montana': (46.595805, -112.027031),
            'Nebraska': (40.809868, -96.675345),
            'Nevada': (39.161921, -119.767),
            'New Hampshire': (43.220093, -71.549896),
            'New Jersey': (40.221741, -74.756138),
            'New Mexico': (35.667231, -105.964575),
            'New York': (42.659829, -73.781339),
            'North Carolina': (35.771, -78.638),
            'North Dakota': (46.813343, -100.779004),
            'Ohio': (39.961176, -82.998794),
            'Oklahoma': (35.482309, -97.534994),
            'Oregon': (44.931109, -123.029159),
            'Pennsylvania': (40.269789, -76.875613),
            'Rhode Island': (41.82355, -71.422132),
            'South Carolina': (34.000, -81.035),
            'South Dakota': (44.367966, -100.336378),
            'Tennessee': (36.165, -86.784),
            'Texas': (30.266667, -97.75),
            'Utah': (40.777477, -111.888237),
            'Vermont': (44.26639, -72.576),
            'Virginia': (37.54, -77.46),
            'Washington': (47.042418, -122.893077),
            'West Virginia': (38.349497, -81.633294),
            'Wisconsin': (43.074722, -89.384444),
            'Wyoming': (41.145548, -104.802042),
            'Washington D.C.': (38.9072, -77.0369),
        }
        
        # Saint Lucia coordinates (Castries)
        self.saint_lucia = (14.0101, -60.9875)
    
    def calculate_tsn(self, temp_c, humidity, wind_speed_ms):
        """
        Calculate Thermal Sensation Number (TSN) using ASHRAE methodology
        This is a simplified approximation of the complex ASHRAE calculation
        
        Parameters:
        temp_c: Temperature in Celsius
        humidity: Relative humidity (0-100)
        wind_speed_ms: Wind speed in m/s
        
        Returns:
        TSN value on ASHRAE 9-point scale (-4 to +4)
        """
        # Convert temperature to apparent temperature considering humidity and wind
        # Heat Index calculation (simplified)
        if temp_c >= 27:  # Above 80°F
            temp_f = temp_c * 9/5 + 32
            rh = humidity
            
            # Simplified heat index formula
            hi = (-42.379 + 2.04901523*temp_f + 10.14333127*rh 
                  - 0.22475541*temp_f*rh - 6.83783e-3*temp_f**2 
                  - 5.481717e-2*rh**2 + 1.22874e-3*temp_f**2*rh 
                  + 8.5282e-4*temp_f*rh**2 - 1.99e-6*temp_f**2*rh**2)
            
            apparent_temp_c = (hi - 32) * 5/9
        else:
            # Wind chill effect for lower temperatures
            if wind_speed_ms > 0:
                wind_speed_kmh = wind_speed_ms * 3.6
                if temp_c <= 10 and wind_speed_kmh >= 4.8:
                    wind_chill = (13.12 + 0.6215*temp_c - 11.37*(wind_speed_kmh**0.16) 
                                 + 0.3965*temp_c*(wind_speed_kmh**0.16))
                    apparent_temp_c = wind_chill
                else:
                    apparent_temp_c = temp_c
            else:
                apparent_temp_c = temp_c
        
        # Convert apparent temperature to TSN scale
        # ASHRAE comfort zone is roughly 20-26°C
        if apparent_temp_c < 5:
            tsn = -4  # Very cold
        elif apparent_temp_c < 10:
            tsn = -3  # Cold
        elif apparent_temp_c < 15:
            tsn = -2  # Cool
        elif apparent_temp_c < 18:
            tsn = -1  # Slightly cool
        elif 18 <= apparent_temp_c <= 26:
            tsn = 0   # Neutral (comfortable)
        elif apparent_temp_c <= 29:
            tsn = 1   # Slightly warm
        elif apparent_temp_c <= 32:
            tsn = 2   # Warm
        elif apparent_temp_c <= 35:
            tsn = 3   # Hot
        else:
            tsn = 4   # Very hot
            
        return tsn
    
    def calculate_cit(self, tsn):
        """
        Calculate Climate Index for Tourism (CIT)
        CIT = 6.4 + 0.4*TSN - 0.281*TSN²
        
        Parameters:
        tsn: Thermal Sensation Number
        
        Returns:
        CIT value
        """
        cit = 6.4 + 0.4*tsn - 0.281*(tsn**2)
        return cit
    
    def get_weather_data_openweather(self, lat, lon, api_key):
        """
        Get weather data from OpenWeatherMap API
        You need to sign up for a free API key at openweathermap.org
        """
        try:
            # Current weather
            url = f"http://api.openweathermap.org/data/2.5/weather?lat={lat}&lon={lon}&appid={api_key}&units=metric"
            response = requests.get(url)
            
            if response.status_code == 200:
                data = response.json()
                return {
                    'temperature': data['main']['temp'],
                    'humidity': data['main']['humidity'],
                    'wind_speed': data['wind']['speed'],
                    'description': data['weather'][0]['description']
                }
            else:
                print(f"Error fetching weather data: {response.status_code}")
                return None
        except Exception as e:
            print(f"Error: {e}")
            return None
    
    def get_annual_climate_data(self, location_name, lat, lon):
        """
        Get annual average climate data based on historical climate normals
        This uses realistic annual averages based on climate data
        """
        # Annual climate averages based on historical data (1991-2020 climate normals)
        climate_data = {
            'Saint Lucia': {
                'temperature': 27.8,  # Annual average temperature
                'humidity': 78.0,     # Annual average relative humidity
                'wind_speed': 4.2     # Annual average wind speed
            },
            # Southeastern States (Warm, Humid)
            'Alabama': {'temperature': 17.8, 'humidity': 68, 'wind_speed': 2.8},
            'Florida': {'temperature': 22.8, 'humidity': 74, 'wind_speed': 3.1},
            'Georgia': {'temperature': 17.3, 'humidity': 68, 'wind_speed': 2.6},
            'Louisiana': {'temperature': 20.1, 'humidity': 75, 'wind_speed': 2.9},
            'Mississippi': {'temperature': 18.4, 'humidity': 71, 'wind_speed': 2.7},
            'South Carolina': {'temperature': 18.2, 'humidity': 69, 'wind_speed': 2.8},
            'Arkansas': {'temperature': 16.8, 'humidity': 68, 'wind_speed': 3.2},
            'Tennessee': {'temperature': 15.1, 'humidity': 65, 'wind_speed': 2.9},
            'North Carolina': {'temperature': 15.8, 'humidity': 65, 'wind_speed': 3.0},
            
            # Southwestern States (Hot, Dry)
            'Arizona': {'temperature': 19.1, 'humidity': 38, 'wind_speed': 2.8},
            'Nevada': {'temperature': 12.9, 'humidity': 42, 'wind_speed': 3.4},
            'New Mexico': {'temperature': 13.7, 'humidity': 45, 'wind_speed': 4.1},
            'Utah': {'temperature': 11.2, 'humidity': 55, 'wind_speed': 3.6},
            
            # California (Mediterranean/Desert)
            'California': {'temperature': 16.3, 'humidity': 62, 'wind_speed': 3.2},
            
            # Texas (Varied)
            'Texas': {'temperature': 19.4, 'humidity': 64, 'wind_speed': 4.2},
            'Oklahoma': {'temperature': 16.2, 'humidity': 62, 'wind_speed': 4.8},
            
            # Pacific Northwest (Mild, Wet)
            'Washington': {'temperature': 10.8, 'humidity': 75, 'wind_speed': 3.1},
            'Oregon': {'temperature': 11.9, 'humidity': 71, 'wind_speed': 2.9},
            
            # Northern States (Cold)
            'Alaska': {'temperature': -2.8, 'humidity': 68, 'wind_speed': 4.2},
            'Montana': {'temperature': 7.1, 'humidity': 58, 'wind_speed': 4.5},
            'North Dakota': {'temperature': 6.4, 'humidity': 65, 'wind_speed': 5.2},
            'South Dakota': {'temperature': 9.1, 'humidity': 64, 'wind_speed': 4.8},
            'Minnesota': {'temperature': 7.7, 'humidity': 67, 'wind_speed': 4.1},
            'Wisconsin': {'temperature': 8.6, 'humidity': 69, 'wind_speed': 3.8},
            'Michigan': {'temperature': 9.2, 'humidity': 71, 'wind_speed': 3.9},
            'Maine': {'temperature': 7.8, 'humidity': 70, 'wind_speed': 3.5},
            'Vermont': {'temperature': 7.6, 'humidity': 68, 'wind_speed': 3.2},
            'New Hampshire': {'temperature': 8.3, 'humidity': 66, 'wind_speed': 2.9},
            
            # Great Lakes/Northeast (Cool, Moderate)
            'New York': {'temperature': 10.1, 'humidity': 65, 'wind_speed': 3.4},
            'Pennsylvania': {'temperature': 10.9, 'humidity': 64, 'wind_speed': 3.1},
            'Ohio': {'temperature': 11.7, 'humidity': 67, 'wind_speed': 3.5},
            'Indiana': {'temperature': 12.1, 'humidity': 67, 'wind_speed': 3.7},
            'Illinois': {'temperature': 11.8, 'humidity': 66, 'wind_speed': 4.0},
            'Iowa': {'temperature': 10.2, 'humidity': 69, 'wind_speed': 4.3},
            'Missouri': {'temperature': 13.9, 'humidity': 65, 'wind_speed': 3.8},
            
            # Mid-Atlantic (Moderate)
            'Virginia': {'temperature': 14.1, 'humidity': 64, 'wind_speed': 2.8},
            'West Virginia': {'temperature': 11.8, 'humidity': 66, 'wind_speed': 2.6},
            'Maryland': {'temperature': 13.2, 'humidity': 63, 'wind_speed': 3.1},
            'Delaware': {'temperature': 13.8, 'humidity': 64, 'wind_speed': 3.4},
            'New Jersey': {'temperature': 12.1, 'humidity': 62, 'wind_speed': 3.6},
            
            # New England (Cool)
            'Massachusetts': {'temperature': 9.7, 'humidity': 64, 'wind_speed': 4.1},
            'Connecticut': {'temperature': 10.8, 'humidity': 63, 'wind_speed': 3.2},
            'Rhode Island': {'temperature': 11.2, 'humidity': 65, 'wind_speed': 4.3},
            
            # Mountain West (Cool, Dry)
            'Colorado': {'temperature': 8.9, 'humidity': 52, 'wind_speed': 3.8},
            'Wyoming': {'temperature': 6.8, 'humidity': 55, 'wind_speed': 4.9},
            'Idaho': {'temperature': 9.1, 'humidity': 61, 'wind_speed': 3.2},
            
            # Great Plains (Variable)
            'Kansas': {'temperature': 13.6, 'humidity': 62, 'wind_speed': 5.1},
            'Nebraska': {'temperature': 10.4, 'humidity': 64, 'wind_speed': 4.7},
            
            # Tropical/Subtropical Islands
            'Hawaii': {'temperature': 24.8, 'humidity': 65, 'wind_speed': 6.2},
            
            # Kentucky
            'Kentucky': {'temperature': 13.8, 'humidity': 66, 'wind_speed': 2.9},

            # Washington D.C. (Mid-Atlantic climate)
            'Washington D.C.': {'temperature': 14.2, 'humidity': 62, 'wind_speed': 3.3}
        }
        
        if location_name in climate_data:
            return climate_data[location_name]
        else:
            # Default temperate climate if not found
            return {'temperature': 12.0, 'humidity': 65, 'wind_speed': 3.5}
    
    def calculate_cit_for_locations(self, data_type='annual_average', api_key=None):
        """
        Calculate CIT for Saint Lucia and all US states
        
        Parameters:
        data_type: 'annual_average' (recommended), 'current_weather', or 'sample'
        api_key: OpenWeatherMap API key (if using current weather data)
        """
        results = []
        
        # Add Saint Lucia
        locations = {'Saint Lucia': self.saint_lucia}
        locations.update(self.us_states)
        
        print("Calculating Climate Index for Tourism (CIT)...")
        print(f"Data Type: {data_type}")
        print("="*60)
        
        for location, coords in locations.items():
            lat, lon = coords
            
            # Get climate data based on selected type
            if data_type == 'annual_average':
                weather = self.get_annual_climate_data(location, lat, lon)
                data_source = "Annual Average"
            elif data_type == 'current_weather' and api_key:
                weather = self.get_weather_data_openweather(lat, lon, api_key)
                if weather is None:
                    weather = self.get_annual_climate_data(location, lat, lon)
                    data_source = "Annual Average (API failed)"
                else:
                    data_source = "Current Weather"
            else:
                # Fallback to annual average
                weather = self.get_annual_climate_data(location, lat, lon)
                data_source = "Annual Average"
            
            # Calculate TSN and CIT
            tsn = self.calculate_tsn(
                weather['temperature'], 
                weather['humidity'], 
                weather['wind_speed']
            )
            cit = self.calculate_cit(tsn)
            
            results.append({
                'Location': location,
                'Latitude': lat,
                'Longitude': lon,
                'Temperature_C': weather['temperature'],
                'Humidity_%': weather['humidity'],
                'Wind_Speed_ms': weather['wind_speed'],
                'TSN': tsn,
                'CIT': round(cit, 2),
                'Data_Source': data_source
            })
            
            print(f"{location:20} | Temp: {weather['temperature']:5.1f}°C | "
                  f"Humidity: {weather['humidity']:5.1f}% | "
                  f"Wind: {weather['wind_speed']:4.1f}m/s | "
                  f"TSN: {tsn:2d} | CIT: {cit:5.2f}")
            
            # Small delay to be respectful to APIs
            if data_type == 'current_weather':
                time.sleep(0.1)
        
        return pd.DataFrame(results)
    
    def interpret_cit(self, cit_value):
        """
        Interpret CIT values for tourism suitability
        Higher CIT values indicate better conditions for tourism
        """
        if cit_value >= 7:
            return "Excellent"
        elif cit_value >= 6:
            return "Very Good"
        elif cit_value >= 5:
            return "Good"
        elif cit_value >= 4:
            return "Fair"
        elif cit_value >= 3:
            return "Poor"
        else:
            return "Very Poor"
    
    def analyze_results(self, df):
        """
        Analyze and summarize CIT results
        """
        print("\n" + "="*50)
        print("CLIMATE INDEX FOR TOURISM (CIT) ANALYSIS")
        print("="*50)
        
        # Add interpretation
        df['Tourism_Suitability'] = df['CIT'].apply(self.interpret_cit)
        
        # Summary statistics
        print(f"\nSummary Statistics:")
        print(f"Mean CIT: {df['CIT'].mean():.2f}")
        print(f"Median CIT: {df['CIT'].median():.2f}")
        print(f"Standard Deviation: {df['CIT'].std():.2f}")
        print(f"Range: {df['CIT'].min():.2f} - {df['CIT'].max():.2f}")
        
        # Top and bottom performers
        print(f"\nTop 5 Locations (Best Climate for Tourism):")
        top_5 = df.nlargest(5, 'CIT')[['Location', 'CIT', 'Tourism_Suitability']]
        for _, row in top_5.iterrows():
            print(f"  {row['Location']:20} | CIT: {row['CIT']:5.2f} | {row['Tourism_Suitability']}")
        
        print(f"\nBottom 5 Locations (Challenging Climate for Tourism):")
        bottom_5 = df.nsmallest(5, 'CIT')[['Location', 'CIT', 'Tourism_Suitability']]
        for _, row in bottom_5.iterrows():
            print(f"  {row['Location']:20} | CIT: {row['CIT']:5.2f} | {row['Tourism_Suitability']}")
        
        # Saint Lucia specific analysis
        st_lucia = df[df['Location'] == 'Saint Lucia'].iloc[0]
        print(f"\nSaint Lucia Analysis:")
        print(f"  CIT Score: {st_lucia['CIT']:.2f}")
        print(f"  Tourism Suitability: {st_lucia['Tourism_Suitability']}")
        print(f"  Rank: {df[df['CIT'] >= st_lucia['CIT']].shape[0]} out of {len(df)}")
        
        return df

# Example usage
def main():
    """
    Main function to demonstrate CIT calculation
    """
    calculator = CITCalculator()
    
    print("Climate Index for Tourism (CIT) Calculator")
    print("Based on de Freitas, Scott, and McBoyle (2008)")
    print("Formula: CIT = 6.4 + 0.4*TSN - 0.281*TSN²")
    print("\nUsing annual climate averages (recommended for tourism analysis)")
    
    # Calculate CIT using annual averages (recommended approach)
    results_df = calculator.calculate_cit_for_locations(data_type='annual_average')
    
    # Analyze results
    final_df = calculator.analyze_results(results_df)
    
    # Save results
    output_file = 'cit_annual_averages_2024.csv'
    final_df.to_csv(output_file, index=False)
    print(f"\nResults saved to {output_file}")
    
    
    return final_df

if __name__ == "__main__":
    # Run the calculation
    df = main()
    
    # Display sample of results
    print("\nSample of Results:")
    print(df.head(10).to_string(index=False))

# Alternative usage examples:
# 
# 1. Using annual climate averages (RECOMMENDED for tourism research):
calculator = CITCalculator()
results = calculator.calculate_cit_for_locations(data_type='annual_average')
# 
# 2. Using current weather data (requires API key):
# API_KEY = "your_openweathermap_api_key_here"
# results = calculator.calculate_cit_for_locations(data_type='current_weather', api_key=API_KEY)
# 
# 3. For seasonal analysis, you could extend this to calculate monthly averages:
# for month in range(1, 13):
#     monthly_results = calculator.calculate_monthly_cit(month)