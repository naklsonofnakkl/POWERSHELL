##This Script will restart Global Protect and its service

# Close the application
Stop-Process -Name "PanGPA"
# Restart the PanGPS service
Restart-Service -Name "PanGPS"
exit