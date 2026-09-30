import heapq
import json

from flask import Flask, jsonify, render_template, request

import travel_data
import country_converter

hotel_countries = country_converter.CountryConverter().data['name_short'].to_list()

app = Flask(__name__)


@app.get("/")
def home():
    return render_template("index.html", hotel_countries=hotel_countries)


@app.get("/locations")
def location_lookup():
    mode = request.args.get("mode", "flight")
    query = request.args.get("q", "").strip().casefold()

    if len(query) < 2:
        return jsonify([])

    matches = (
        [
            {
                'code': iata,
                'name': airport['name'],
                'countryCode': airport['country'],
                'latitude': airport['lat'],
                'longitude': airport['lon']
            }
            for iata, airport in travel_data.airports.items()
            if (
                query in iata.casefold()
                or query in airport['name'].casefold()
                or query in airport['city'].casefold()
                or query in airport['country'].casefold()
        )
        ]
        if mode == 'flight' else
        travel_data.eu_stations[
            travel_data.eu_stations['is_suggestable']
            & (travel_data.eu_stations['name'].str.contains(query, case=False)
            | travel_data.eu_stations['slug'].str.contains(query, case=False)
            | travel_data.eu_stations['info:en'].str.contains(query, case=False)
            | travel_data.eu_stations['code'].str.contains(query, case=False)
            )]
        .rename(columns={
            'country': 'countryCode',
            # 'atoc_id': 'code',
        })
        .to_dict('records')
    )
    # Put exact code matches first, followed by names alphabetically.
    # For rail, put GB stations first
    matches = heapq.nsmallest(10, matches,
        key=lambda location: (
            location["code"] != query.upper(),
            not (mode == 'rail' and location['countryCode'] == 'GB'),
            location["name"],
        )
    )

    # Country code lookup now (otherwise it's slow!)
    for match in matches:
        match['country'] = country_converter.convert(match['countryCode'], to='name')
        match['countryFlag'] = "".join(chr(ord(char) + 127397) for char in match['countryCode'])

    # Remove NaNs
    matches = [{key: value for key, value in match.items() if value == value} for match in matches]

    return jsonify(matches)


@app.post("/calculate")
def calculate():
    mode = request.form["mode"]
    locations = json.loads(request.form.get("locations", "[]"))
    include_return = request.form.get("return_same") == "on"

    hotel_country = request.form.get("hotel_country", "").strip()
    try:
        hotel_nights = max(0, int(request.form.get("hotel_nights", 0)))
    except ValueError:
        hotel_nights = 0

    if len(locations) < 2:
        return "At least two journey locations are required.", 400

    trip = {'origin': locations[0], 'destination': locations[-1]}
    if include_return:
        locations.extend(locations[-2:-1])

    trip |= {
        "mode": mode,
        "locations": locations,
        "hotel_country": hotel_country,
        "hotel_nights": hotel_nights,
    }

    hotel_rate = travel_data.hotel_conversions[hotel_country if hotel_country in travel_data.hotel_conversions else None]
    result = {'hotel': hotel_rate * hotel_nights}

    alternative = False
    polyline = ''
    if mode == "flight":
        total_distance, hauls, emissions = travel_data.flight_distance_from_codes(
            [location['code'] for location in locations])
        alternative = trip['destination']['countryCode'] in travel_data.western_europe
        if alternative:
            _, alt_emissions, polyline = travel_data.total_train_distance_from_coords([
                (location['latitude'], location['longitude'], location['code'])
                for location in locations
            ])
        coord_pairs = [(location['latitude'], location['longitude']) for location in locations]
    else:  # rail
        total_distance, emissions, polyline = travel_data.total_train_distance_from_coords([
            (location['latitude'], location['longitude'], location['code'])
            for location in locations
        ])
        coord_pairs = []
    trip['total_distance'] = round(total_distance)
    result['travel'] = emissions
    result['total'] = round(result['hotel'] + result['travel'])
    result['hotel'] = round(result['hotel'])
    result['travel'] = round(result['travel'])
    if alternative and alt_emissions < emissions:
        result['alt_total'] = round(result['hotel'] + alt_emissions)
        result['reduction_percent'] = round(100 * (1 - result['alt_total'] / result['total']))

    # result = {
    #     "travel": 145.2,
    #     "hotel": 22.7,
    #     "total": 167.9,
    # }

    return render_template(
        "results.html",
        trip=trip,
        result=result,
        rail_polyline=polyline,
        flight_pairs=coord_pairs,
    )


if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0')
