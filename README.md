# User Modal Web Part for SharePoint

A modern SharePoint Framework (SPFx) web part that displays team members in a tile-based layout with modal functionality for detailed information.

## Features

- Displays 1-4 user tiles per view with configurable layout
- Carousel navigation for browsing through additional users
- Pulls content dynamically from a SharePoint list with user information
- Modal window shows detailed user information when a tile is clicked
- Integrates with SharePoint user profiles to get profile photos and job titles
- Responsive design that works across all device sizes
- Modern UI with shadows, rounded corners, and hover effects
- Configurable fields for customization

![User Modal Web Part](./assets/user-modal-preview.png)

## Getting Started

### Prerequisites

- Node.js (version 18.17.1 or higher)
- SharePoint Developer environment
- SharePoint list with the correct content type (see below)

### Installation

1. Clone this repository
2. Run `npm install`
3. Run `gulp serve` to test locally
4. Run `gulp bundle --ship` and `gulp package-solution --ship` to package for deployment
5. Upload the `.sppkg` file from the `sharepoint/solution` folder to your SharePoint App Catalog
6. Add the web part to your page

## SharePoint List Setup

### Content Type Columns

Create a SharePoint list with the following columns:

1. **Title** (Default column)
   - Used for the list item title, not displayed in the web part

2. **User**
   - Type: Person or Group
   - Allow multiple selections: No
   - Show field: User name
   - This field connects to the SharePoint user profile

3. **Description**
   - Type: Multiple lines of text
   - Used for the user's detailed description in the modal

4. **Certification**
   - Type: Multiple lines of text
   - Used to list user certifications in the modal

### Creating the List

1. Create a new list in SharePoint (suggested name: "TeamMembers")
2. Add the columns specified above
3. Add your team members as list items
4. Configure the web part to use this list

## Web Part Configuration

In the web part properties pane, you can configure:

- **Web Part Title**: The heading displayed above the tiles
- **SharePoint List Name**: Name of the list containing your user information
- **Tiles Per View**: Number of tiles to display at once (1-4)
- **Field Name Settings**: Configure custom field names if they differ from defaults

## Technical Details

### PnP JS Integration

This web part uses PnP JS (SharePoint Patterns and Practices JavaScript library) for SharePoint data operations, which offers several advantages:

- Cleaner, more maintainable code for SharePoint operations
- Improved error handling and fallback mechanisms
- Better performance through optimized queries
- Access to SharePoint user profiles for profile photos and job title information

### User Profile Integration

The web part implements a multi-layered approach to retrieve user data:

1. **Basic Information**: Gets user name and email from the SharePoint list
2. **Profile Enhancement**: Retrieves additional profile information such as job title
3. **Profile Photo**: Attempts to get the user's profile photo, with fallback to a default image

### Modal Implementation

The modal dialog is implemented using Fluent UI components:

1. **Responsive Design**: Works well on all screen sizes
2. **Accessible**: Implements proper keyboard navigation and accessibility features
3. **Themed**: Supports SharePoint themes including dark mode

## Development Notes

- Built using SharePoint Framework (SPFx) 1.20.0
- Uses React and Fluent UI components
- Implements responsive grid layout with CSS Grid
- Leverages PnP JS library for enhanced SharePoint operations
