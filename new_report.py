import pandas as pd
import numpy as np

# Read the CSV file
def calculate_disk_conversion_savings():
    # Read the CSV file using the window.fs API
    try:
        file_content = window.fs.readFile('SSD Persistand Disk Data (cas-prod-env) - Sheet1.csv', { encoding: 'utf8' })
        df = pd.read_csv(pd.StringIO(file_content), skipinitialspace=True)
    except Exception as e:
        print(f"Error reading file: {e}")
        return None
    
    # Clean column names by stripping whitespace
    df.columns = [col.strip() for col in df.columns]
    
    # Calculate costs and savings
    # SSD persistent disk pricing ~$0.17 per GB per month
    # Balanced persistent disk pricing ~$0.10 per GB per month (41.2% cheaper)
    ssd_cost_per_gb = 0.17
    balanced_cost_per_gb = 0.10
    savings_percentage = 41.2  # 41.2% savings
    
    # Initialize lists to store data for our output table
    disk_data = []
    
    # Process each row in the dataframe
    for _, row in df.iterrows():
        # Process boot disk
        if row['Boot Disk Size (GB)'] != 'No Record':
            boot_disk_size = float(row['Boot Disk Size (GB)'])
            vm_name = row['VM Name']
            attached_to = row['VM Name']
            zone = row['Zone']
            region = zone.split('-')[0] + '-' + zone.split('-')[1]
            is_boot_disk = 'Yes'
            
            # Calculate costs
            current_monthly_cost = round(boot_disk_size * ssd_cost_per_gb, 1)
            balanced_monthly_cost = round(boot_disk_size * balanced_cost_per_gb, 1)
            monthly_savings = round(current_monthly_cost - balanced_monthly_cost, 1)
            annual_savings = round(monthly_savings * 12, 1)
            
            disk_data.append({
                'Disk Name': row['Boot Disk Name'],
                'Disk Size (GB)': boot_disk_size,
                'Is Boot Disk': is_boot_disk,
                'Attached To': attached_to,
                'Zone': zone,
                'Region': region,
                'Current Monthly Cost (USD)': current_monthly_cost,
                'Balanced Monthly Cost (USD)': balanced_monthly_cost,
                'Monthly Savings (USD)': monthly_savings,
                'Annual Savings (USD)': annual_savings,
                'Savings Percentage': savings_percentage
            })
        
        # Process additional disk if present
        if row['Additional Disk Name'] != 'No Record':
            additional_disk_size = float(row['Additional Disk Size (GB)'])
            is_boot_disk = 'No'
            
            # Calculate costs
            current_monthly_cost = round(additional_disk_size * ssd_cost_per_gb, 1)
            balanced_monthly_cost = round(additional_disk_size * balanced_cost_per_gb, 1)
            monthly_savings = round(current_monthly_cost - balanced_monthly_cost, 1)
            annual_savings = round(monthly_savings * 12, 1)
            
            disk_data.append({
                'Disk Name': row['Additional Disk Name'],
                'Disk Size (GB)': additional_disk_size,
                'Is Boot Disk': is_boot_disk,
                'Attached To': row['VM Name'],
                'Zone': row['Zone'],
                'Region': zone.split('-')[0] + '-' + zone.split('-')[1],
                'Current Monthly Cost (USD)': current_monthly_cost,
                'Balanced Monthly Cost (USD)': balanced_monthly_cost,
                'Monthly Savings (USD)': monthly_savings,
                'Annual Savings (USD)': annual_savings,
                'Savings Percentage': savings_percentage
            })
    
    # Create DataFrame from the collected data
    output_df = pd.DataFrame(disk_data)
    
    # Calculate summary information
    total_disk_count = len(output_df)
    total_disk_size = output_df['Disk Size (GB)'].sum()
    current_monthly_cost = output_df['Current Monthly Cost (USD)'].sum()
    balanced_monthly_cost = output_df['Balanced Monthly Cost (USD)'].sum()
    potential_monthly_savings = output_df['Monthly Savings (USD)'].sum()
    potential_annual_savings = output_df['Annual Savings (USD)'].sum()
    avg_savings_percentage = savings_percentage  # Fixed at 41.2%
    
    # Create summary DataFrame
    summary_data = {
        'Metric': [
            'Total Disk Count:', 
            'Total Disk Size (GB):', 
            'Current Monthly Cost (USD):', 
            'Projected Monthly Cost with Balanced Disks (USD):', 
            'Potential Monthly Savings (USD):', 
            'Potential Annual Savings (USD):', 
            'Average Savings Percentage:'
        ],
        'Value': [
            total_disk_count,
            round(total_disk_size, 0),
            round(current_monthly_cost, 2),
            round(balanced_monthly_cost, 2),
            round(potential_monthly_savings, 2),
            round(potential_annual_savings, 0),
            f"{avg_savings_percentage}%"
        ]
    }
    summary_df = pd.DataFrame(summary_data)
    
    # Save the detailed disk data to a CSV file
    output_df.to_csv('disk_conversion_details.csv', index=False)
    
    # Create Excel file with formatting
    with pd.ExcelWriter('US-West1_SSD_to_Balanced_Conversion_Report.xlsx', engine='openpyxl') as writer:
        # Create empty DataFrame for title
        title_df = pd.DataFrame({' ': ['', '', '']})
        title_df.to_excel(writer, sheet_name='Conversion Report', startrow=0, startcol=4, header=False, index=False)
        
        # Write summary data
        summary_df.to_excel(writer, sheet_name='Conversion Report', startrow=2, startcol=0, header=False, index=False)
        
        # Write disk details
        output_df.to_excel(writer, sheet_name='Conversion Report', startrow=12, startcol=0, index=False)
        
        # Get the workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Conversion Report']
        
        # Add title
        worksheet.cell(row=1, column=5, value='US-West1 (Oregon) SSD to Balanced Disk Conversion Savings Report')
        
        # Format title
        title_cell = worksheet.cell(row=1, column=5)
        title_cell.font = workbook.add_font(bold=True, size=14)
        
        # Format summary section
        worksheet.cell(row=3, column=1, value='Summary')
        summary_cell = worksheet.cell(row=3, column=1)
        summary_cell.font = workbook.add_font(bold=True)
    
    # Return summary information for display
    return {
        'total_disk_count': total_disk_count,
        'total_disk_size': total_disk_size,
        'current_monthly_cost': current_monthly_cost,
        'balanced_monthly_cost': balanced_monthly_cost,
        'potential_monthly_savings': potential_monthly_savings,
        'potential_annual_savings': potential_annual_savings,
        'avg_savings_percentage': avg_savings_percentage
    }

# Call the function to calculate savings
results = calculate_disk_conversion_savings()

# Display results summary
if results:
    print("SSD to Balanced Disk Conversion Analysis Complete")
    print(f"Total Disk Count: {results['total_disk_count']}")
    print(f"Total Disk Size: {results['total_disk_size']} GB")
    print(f"Current Monthly Cost: ${results['current_monthly_cost']:.2f}")
    print(f"Projected Monthly Cost with Balanced Disks: ${results['balanced_monthly_cost']:.2f}")
    print(f"Potential Monthly Savings: ${results['potential_monthly_savings']:.2f}")
    print(f"Potential Annual Savings: ${results['potential_annual_savings']:.0f}")
    print(f"Average Savings Percentage: {results['avg_savings_percentage']}%")
    print("\nFiles created:")
    print("- disk_conversion_details.csv")
    print("- US-West1_SSD_to_Balanced_Conversion_Report.xlsx")
else:
    print("Error processing the data. Please check the input file.")