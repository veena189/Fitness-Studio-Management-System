# Fitness Studio Management System

A small-scale application that applies database management, Excel, and VBA concepts to streamline client bookings, trainer assignments, and revenue tracking for fitness studios. It provides an intuitive interface for efficient operations, ensuring smooth management for business owners and staff.

## Description

The **Fitness Studio Management System** integrates an Access database, Excel front-end, and VBA middleware to manage various tasks in a fitness studio, including:
- Managing client bookings
- Assigning trainers to classes
- Tracking revenue by class type and trainer specialization
- Managing client-specific pending bookings

This system simplifies operational tasks, helping fitness studio owners and staff save time, reduce errors, and improve business management.

## Features

- **Client Management**: Store and manage client details, including bookings and statuses.
- **Trainer Assignment**: Assign trainers to specific classes based on their specialization.
- **Revenue Tracking**: Analyze revenue by class type and trainer specialization.
- **Pending Bookings**: Retrieve and manage pending bookings for clients.
- **Excel Interface**: Use an intuitive Excel front-end with VBA integration for seamless data handling.

## Database Design

The project uses a Microsoft Access database with the following interrelated tables:

1. **Clients Table**: Stores client details (Client ID, Name, Email).
2. **Trainers Table**: Stores trainer details (Trainer ID, Name, Specialization).
3. **ClassTypes Table**: Stores class information (Class Type ID, Name, Trainer, Price, Schedule).
4. **ClientBookings Table**: Tracks client bookings (Booking ID, Client ID, Class Type ID, Date, Status).

### Relationships
- ClassTypes and Trainers are linked via Trainer ID.
- ClientBookings is linked to Clients and ClassTypes via Client ID and ClassType ID, respectively.

## Front-End: Excel Interface

The Excel workbook contains the following worksheets:

1. **Bookings_Dashboard**: Displays all client bookings with color-coded statuses (Confirmed, Pending, Cancelled).
2. **Revenue_Insights**: Uses pivot tables to analyze revenue by class type and trainer specialization.
3. **Add_New_Booking**: Input form for adding new bookings and saving them to the database.
4. **ClientPendingBookings**: Retrieves and displays pending bookings for a specific client.

## VBA Middleware

The VBA middleware links the Excel front-end to the Access database. Key modules include:

- **AllFunctionsUsedModule**: Contains reusable functions and data structures for managing bookings.
- **LoadBookingsModule**: Loads all bookings into the Bookings_Dashboard worksheet.
- **SaveBookingsToDatabaseModule**: Saves new bookings entered by users to the database.
- **PendingClientDetailsModule**: Retrieves pending bookings for a specific client based on their ID.

## Files Description

1. **FitnessStudioManagement.accdb**: This is the Microsoft Access database file that stores all the data related to clients, trainers, class types, and bookings. The database contains four key tables: `Clients`, `Trainers`, `ClassTypes`, and `ClientBookings`, and is used to manage and retrieve data in the system.
   
2. **FitnessStudioManagement.xlsm**: This is the Excel workbook that serves as the front-end for the application. It includes four worksheets for managing bookings, revenue insights, adding new bookings, and viewing pending bookings. It also contains embedded VBA code that interacts with the database to perform various operations such as loading, saving, and querying data.

## Setup Instructions

Follow these steps to set up and use the Fitness Studio Management System:

1. **Download all files** from the repository.
2. Place **FitnessStudioManagement.xlsm** and **FitnessStudioManagement.accdb** in the same folder.
3. Open **FitnessStudioManagement.xlsm** in Microsoft Excel.
4. **Enable macros** to use all the application features.

### Usage Instructions:

- **Load Bookings**: Click the "Load Bookings from Database" button to fetch and display all booking records.
- **Add New Booking**: Enter new booking details in the "Add_New_Booking" sheet and click "Save Bookings to Database" to add it.
- **View Pending Bookings**: Use the "ClientPendingBookings" sheet to check pending bookings for a specific client.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.


## Future Enhancements

- **Additional Features**: Add payment records, membership plans, and attendance tracking.
- **User Authentication**: Implement role-based access controls for secure data handling.
- **Improved Interface**: Enhance the front-end with advanced filtering, better user experience, and interactive dashboards.
- **Automated Notifications**: Integrate email notifications for booking confirmations and cancellations.
- **Cloud Integration**: Host the database in the cloud and transition to a web-based interface for greater scalability.

