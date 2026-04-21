# Microscope-QC-Testing-
**SmartQC** is a desktop quality control app for smart ring manufacturing. It connects to a USB microscope, lets operators capture ring images, and classify them as Accepted, Rejected, or Rework using simple keyboard shortcuts — with automatic daily folder organisation, Excel logging, and database tracking.
Here is your 200 word project description:

---

**SmartRingQC — Smart Ring Quality Control Inspection Software**

SmartRingQC is a professional desktop application built for smart ring manufacturing companies to streamline their quality control inspection process. The software connects directly to a digital microscope via USB and provides operators with a fast, keyboard-driven inspection workflow.

Operators capture ring images by pressing the Spacebar, which instantly saves each photo into a waiting queue. From there, inspectors review each image on screen and classify it in one keystroke — pressing A to Accept, R to Reject, or W to send for Rework. Every decision automatically moves the image into the correct folder, updates the SQLite database, and logs the action into an Excel report in real time.

One of the standout features is the automatic daily folder system. Every day the software creates a fresh dated folder inside each classification category, named in the format 18-April-2026-Friday, keeping inspection data cleanly organized by date for easy traceability and reporting.

The dashboard displays live statistics including total accepted, rejected, rework counts and the daily acceptance rate. Operators can also undo any mistake instantly using the Z key.

Built with Python, OpenCV, Tkinter, and openpyxl, SmartRingQC is lightweight, reliable, and deployable as a standalone desktop application requiring no internet connection.
