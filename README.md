# Belvedere Booking

This repository contains a Google Apps Script used to set up a booking system in Google Sheets. The script creates sheets for properties, buildings, rooms and bookings. A helper function is provided to sync bookings to a Google Calendar.

## Setting Up
1. Open the script in Google Apps Script and run `setupBookingSystem()` to create the initial sheets.
2. Enter your data or use the provided sample data.

## Syncing with Google Calendar
Run the `syncBookingsToGoogleCalendar()` function to push all bookings to a calendar. The script prompts for a calendar ID and creates private events for each booking so that the calendar only shows that a time slot is booked without revealing details.
