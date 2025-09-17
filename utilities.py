import streamlit as st
import folium
from prettymapp.geo import get_aoi
from prettymapp.osm import get_osm_geometries
from prettymapp.plotting import Plot
from prettymapp.settings import STYLES


def create_map():
    # Center on a default location (latitude, longitude) with a base zoom level
    return folium.Map(location=[20, 50], zoom_start=2)


def plot_pretty_map(lat, lon):
    try:
        # Define the area of interest (AOI) around the given latitude and longitude
        aoi = get_aoi(address=f"{lat},{lon}", radius=1000, rectangular=False)
        # Fetch OpenStreetMap geometries (roads, parks, buildings, etc.) within that area
        df = get_osm_geometries(aoi=aoi)
        # Use prettymapp's Plot to draw the map with a chosen style
        fig = Plot(df=df, aoi_bounds=aoi.bounds, draw_settings=STYLES["Citrus"]).plot_all()
        return fig  # fig is a Matplotlib figure
    except Exception as e:
        st.error(f"Error: {e}")  # Display an error message in the Streamlit app
        return None