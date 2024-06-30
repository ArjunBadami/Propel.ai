import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd
from datetime import datetime, timedelta
import matplotlib.gridspec as gridspec
#import excel_parser
import os
import plotly.figure_factory as ff
import plotly.io as pio
import matplotlib.dates as mdates

# Create a toy example data set with additional details
def create_gantt(tasks, output_path):
    num_tasks = len(tasks)
    # Create a mapping from task names to their y-axis positions
    task_y_positions = {}
    for idx, task in enumerate(tasks):
        task_y_positions[task['Task']] = idx
    
    #fig = plt.figure(figsize=(140, 60))
    #gs = gridspec.GridSpec(1, 2, width_ratios=[1, 1], wspace=0.0)

    # Create the Gantt chart subplot
    #ax = fig.add_subplot(gs[1])
    # Define base figure size and scaling factor
    base_width = 10
    base_height_per_task = 0.45  # Adjust this factor to control the height scaling
    # Calculate the figure height based on the number of tasks
    fig_height = max(6, len(tasks) * base_height_per_task)
    # Create the figure and subplots
    fig, ax = plt.subplots(figsize=(base_width, fig_height))
    # Add tasks to the Gantt chart
    for i, task in enumerate(tasks):
        ax.barh(task_y_positions[task['Task']], task['Duration'], left=task['Start'], color='blue', edgecolor='black')

    # Reverse the y-axis to match the order of the table
    ax.invert_yaxis()
    def draw_arrow_between_jobs(fig, first_job_dict, second_job_dict):
        ## retrieve tick text and tick vals
        #job_yaxis_mapping = dict(zip(fig.layout.yaxis.ticktext,fig.layout.yaxis.tickvals))
        jobs_delta = second_job_dict['Start'] - first_job_dict['Finish']
        ## horizontal line segment
        fig.add_shape(
            x0=first_job_dict['Finish'], y0=task_y_positions[first_job_dict['Task']], 
            x1=first_job_dict['Finish'] + jobs_delta/2, y1=task_y_positions[first_job_dict['Task']],
            line=dict(color="blue", width=2)
        )
        ## vertical line segment
        fig.add_shape(
            x0=first_job_dict['Finish'] + jobs_delta/2, y0=task_y_positions[first_job_dict['Task']], 
            x1=first_job_dict['Finish'] + jobs_delta/2, y1=task_y_positions[second_job_dict['Task']],
            line=dict(color="blue", width=2)
        )
        ## horizontal line segment
        fig.add_shape(
            x0=first_job_dict['Finish'] + jobs_delta/2, y0=task_y_positions[second_job_dict['Task']], 
            x1=second_job_dict['Start'], y1=task_y_positions[second_job_dict['Task']],
            line=dict(color="blue", width=2)
        )
        ## draw an arrow
        fig.add_annotation(
            x=second_job_dict['Start'], y=task_y_positions[second_job_dict['Task']],
            xref="x",yref="y",
            showarrow=True,
            ax=-10,
            ay=0,
            arrowwidth=2,
            arrowcolor="blue",
            arrowhead=2,
        )
        return fig
    
    def draw_arrow_between_jobs2(ax, first_job_dict, second_job_dict):
        # Calculate the jobs delta
        jobs_delta = second_job_dict['Start'] - first_job_dict['Finish']
        
        # Horizontal line segment
        ax.plot([first_job_dict['Finish'], first_job_dict['Finish'] + jobs_delta/2], 
                [task_y_positions[first_job_dict['Task']], task_y_positions[first_job_dict['Task']]], 
                color="skyblue", linewidth=2)
        
        # Vertical line segment
        ax.plot([first_job_dict['Finish'] + jobs_delta/2, first_job_dict['Finish'] + jobs_delta/2], 
                [task_y_positions[first_job_dict['Task']], task_y_positions[second_job_dict['Task']]], 
                color="skyblue", linewidth=2)
        
        # Horizontal line segment
        ax.plot([first_job_dict['Finish'] + jobs_delta/2, second_job_dict['Start']], 
                [task_y_positions[second_job_dict['Task']], task_y_positions[second_job_dict['Task']]], 
                color="skyblue", linewidth=2)
        
        
        # Draw an arrow
        ax.annotate(
            '', 
            xy=(second_job_dict['Start'], task_y_positions[second_job_dict['Task']]), 
            xytext=(second_job_dict['Start']- timedelta(days=1), task_y_positions[second_job_dict['Task']]), 
            arrowprops=dict(arrowstyle='->', color='skyblue', lw=2)
        )
        
        return ax
    
    # Add dependencies (arrows)
    for i, task in enumerate(tasks):
        for dependency in task['Dependencies']:
            dep_task = next((t for t in tasks if t['Task'] == dependency), None)
            
            dep_end = dep_task['Finish']
            task_start = task['Start']

            # Adjust y-positions for the inverted axis
            y_end = task_y_positions[task['Task']]
            y_start = task_y_positions[dependency]

            ax = draw_arrow_between_jobs2(ax, dep_task, task)
            
            
    
    #fig.write_image(output_path)
    
    # Formatting the Gantt chart
    ax.xaxis.set_ticks_position('top')
    locator = mdates.AutoDateLocator()
    formatter = mdates.AutoDateFormatter(locator)
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(formatter)
    plt.setp(ax.get_xticklabels(), rotation=45, ha='right')
    #ax.set_xlabel('Date')
    #ax.set_ylabel('Task')
    task_names = [task['Name'] for task in tasks]
    ax.set_yticks(list(task_y_positions.values()))
    ax.set_yticklabels(task_names)
    ax.grid(True)
    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    # Create the table
    '''
    #ax_table = fig.add_subplot(gs[0])
    cell_text = [[task['ID'], task['Name']] for task in tasks]
    columns = ['ID', 'Name']
    table = plt.table(cellText=cell_text, colLabels=columns, cellLoc='center', loc='left', colWidths=[0.05, 0.2])

    # Adjust table scale to match the Gantt chart
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.2, 1.2)
    #ax_table.axis('off')
    # Align the table rows with the Gantt chart rows
    for pos, cell in table.get_celld().items():
        if pos[0] == 0:  # Header row
            cell.set_fontsize(12)
            cell.set_text_props(weight='bold')
        else:  # Data rows
            cell.set_height(1.0 / len(tasks))  # Set the cell height to match the Gantt chart

    # Adjust the layout to fit both subplots nicely
    plt.subplots_adjust(left=0.25, right=0.95, top=0.95, bottom=0.1)
    '''
    #plt.show()
    plt.savefig(output_path)
    


#project_tasks = excel_parser.read_excel_tasks("Sample.xlsx")
#project_tasks = excel_parser.convert_tasks_for_gantt(project_tasks)
#create_gantt(project_tasks, "temp.png")
