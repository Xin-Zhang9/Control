# PCB
import serial
import serial.tools.list_ports
# Microscope
import os
import re
import czifile
import bioformats
import javabridge
import numpy as np
import win32com.client
import matplotlib.pyplot as plt
import ReadAnalyzeImageData as rad
# Image analysis
import cv2
import numpy as np
import os
import csv
import matplotlib.pyplot as plt
# RBF
from IPython.display import clear_output
from mpl_toolkits.mplot3d import Axes3D
from scipy.signal import find_peaks
import matplotlib.ticker as ticker
from itertools import zip_longest
import matplotlib.dates as mdates
from mpl_toolkits import mplot3d
import matplotlib.pyplot as plt
from datetime import datetime
import numpy as np
import datetime
import pyvisa
import serial
import time
import csv
import os 


"""______________________________PCB USB port defination_______________________________"""
def detect_ports():
    """Detect available serial ports."""
    ports = serial.tools.list_ports.comports()
    available_ports = [port.device for port in ports]
    print("Available Ports:")
    for port in available_ports:
        print(f" - {port}")
    return available_ports

def connect_to_port(port_name, baud_rate=9600, timeout=1):
    """Connect to a specific serial port."""
    try:
        ser = serial.Serial(port_name, baudrate=baud_rate, timeout=timeout)
        print(f"Connected to {port_name} at {baud_rate} baud.")
        return ser
    except serial.SerialException as e:
        print(f"Error connecting to port {port_name}: {e}")
        return None

def send_ascii_data(ser, data):
    """Send ASCII data through the serial port."""
    if ser and ser.is_open:
        ser.write(data.encode('ascii'))  # Convert string to bytes
        print(f"Sent: {data}")
    else:
        print("Serial port is not open.")

def receive_data(ser, bytes_to_read=100):
    """Receive data from the serial port."""
    if ser and ser.is_open:
        received_data = ser.read(bytes_to_read).decode('ascii')  # Convert bytes to string
        print(f"Received: {received_data}")
        return received_data
    else:
        print("Serial port is not open.")
        return ""

def close_port(ser):
    """Close the serial port."""
    if ser and ser.is_open:
        ser.close()
        print("Serial port closed.")
    else:
        print("Serial port is already closed.")

"""___________________________PCB connect and initialization___________________________"""
# Connect COM port
if __name__ == "__main__":
    # Step 1: Detect available ports
    ports = detect_ports()
    if not ports:
        print("No available ports found.")
    else:
        # Step 2: Connect to the last available port
        port_name = ports[0]  # Replace with the desired port
        print(port_name)
        ser = connect_to_port(port_name, baud_rate=115200)
send_ascii_data(ser, "0\n")
response = receive_data(ser, bytes_to_read=10000)
send_ascii_data(ser, "2\n")
response = receive_data(ser, bytes_to_read=10000)
send_ascii_data(ser, "2\n")
response = receive_data(ser, bytes_to_read=10000)

# Initialize the VISA resource manager
rm = pyvisa.ResourceManager()

# # Open a connection to the electrometer
# # Replace 'GPIB0::24::INSTR' with the actual resource name of your instrument
# electrometer = rm.open_resource('USB0::0x2A8D::0xBA01::MY61100112::0::INSTR')

"""______________________________Zen Microscope Defination_____________________________"""
# Initialize ZEN
Zen = win32com.client.GetActiveObject("Zeiss.Micro.Scripting.ZenWrapperLM")

# Define the folder to save images
save_folder = "C:/Users/PingLab_PC8/Desktop/XinZhang/Zen_python_control/1/"

# Ensure the save folder exists
os.makedirs(save_folder, exist_ok=True)

# def get_next_image_number(folder, prefix="t", extension=".jpeg"):
#     # List all files in the folder
#     files = os.listdir(folder)
#     max_number = 0

#     # Regex to match filenames like t001.jpeg, t002.jpeg
#     pattern = re.compile(rf"{prefix}(\d+){extension}")

#     for file in files:
#         match = pattern.match(file)
#         if match:
#             num = int(match.group(1))  # Extract the number
#             max_number = max(max_number, num)

#     return max_number + 1  # Return the next number

# def Zen_capture():
#     # Get the next starting number
#     start_number = get_next_image_number(save_folder)

#     # Acquire the image
#     image = Zen.Acquisition.AcquireImage()

#     # Generate the filename in the format t001, t002, etc.
#     image_filename = f"{save_folder}t{start_number:03d}.jpeg"  # Format as t001, t002, ...

#     # Save the image
#     image.Save_2(image_filename)

#     return image_filename
def Zen_capture():
    experiment_name = "Xin_Cell_Ca+_every3s_10%_50ms_AI" 
    experiment = Zen.Acquisition.Experiments.GetByName(experiment_name)
    image = Zen.Acquisition.Execute(experiment)  # Run the experiment
    # Zen.Application.Documents.Add(image)
    # Save file from Zen to my folder
    image_filename = save_folder + image.Name
    image.Save_2(image_filename)
    # print(filename)


    return image_filename

"""________________________________Data analyze Defination_____________________________"""
def get_YM_value(image_filename):

    # Load the CZI file
    with czifile.CziFile(image_filename) as czi:
        # Extract the 6D image data as a NumPy array
        image_data = czi.asarray()

    # Print the shape of the loaded image
    # Dimensions: (S, T, Z, C, Y, X)
    # print(f"Image data shape: {image_data.shape}")

    # Define the region of interest (ROI)
    x_start = 1243  # Starting x-coordinate
    y_start = 499  # Starting y-coordinate
    roi_width = 70
    roi_height = 70

    scene = 0           # Scene index
    channel_index = 0   # Choose the channel (e.g., 0 or 1)

    # Extract the corresponding 2D image slice
    image_slice = image_data[scene, channel_index, :, :, 0]  # Full Y and X for the channel

    # Extract the ROI (200x200 starting at x=100, y=100)
    roi = image_slice[y_start:y_start+roi_height, x_start:x_start+roi_width]

    # Verify the shape of the ROI
    # print(f"Shape of ROI: {roi.shape}")  # Should be (200, 200)

    # Calculate mean intensity for the ROI
    mean_intensity = np.mean(roi)
    # print(f"Mean intensity of ROI: {mean_intensity}")

    # Set vmin and vmax to improve contrast
    vmin = np.min(roi)  # Minimum pixel intensity
    vmax = np.max(roi)  # Maximum pixel intensity
    # print(f"Min intensity: {vmin}, Max intensity: {vmax}")


    # # Display the ROI
    # plt.imshow(roi, cmap="gray", vmin=vmin, vmax=vmax)
    # plt.title(f"ROI: x={x_start}, y={y_start}, width={roi_width}, height={roi_height}")
    # plt.colorbar(label="Pixel Intensity")
    # ROI_output_filename = f'{image_filename}_ROI_Intensity.png'
    # output_filepath = os.path.join(save_folder, ROI_output_filename)
    # plt.savefig(output_filepath, dpi=300)
    # plt.close()
    return mean_intensity
    
def get_Background_value(image_filename):

    # Load the CZI file
    with czifile.CziFile(image_filename) as czi:
        # Extract the 6D image data as a NumPy array
        image_data = czi.asarray()

    # Print the shape of the loaded image
    # Dimensions: (S, T, Z, C, Y, X)
    # print(f"Image data shape: {image_data.shape}")

    # Define the region of interest (ROI)
    x_start = 185  # Starting x-coordinate
    y_start = 1122  # Starting y-coordinate
    roi_width = 70
    roi_height = 70


    scene = 0           # Scene index
    channel_index = 0   # Choose the channel (e.g., 0 or 1)

    # Extract the corresponding 2D image slice
    image_slice = image_data[scene, channel_index, :, :, 0]  # Full Y and X for the channel

    # Extract the ROI (200x200 starting at x=100, y=100)
    roi = image_slice[y_start:y_start+roi_height, x_start:x_start+roi_width]

    # Verify the shape of the ROI
    # print(f"Shape of ROI: {roi.shape}")  # Should be (200, 200)

    # Calculate mean intensity for the ROI
    mean_intensity = np.mean(roi)
    # print(f"Mean intensity of ROI: {mean_intensity}")

    # Set vmin and vmax to improve contrast
    vmin = np.min(roi)  # Minimum pixel intensity
    vmax = np.max(roi)  # Maximum pixel intensity
    # print(f"Min intensity: {vmin}, Max intensity: {vmax}")

    # # Display the ROI
    # plt.imshow(roi, cmap="gray", vmin=vmin, vmax=vmax)
    # plt.title(f"ROI: x={x_start}, y={y_start}, width={roi_width}, height={roi_height}")
    # plt.colorbar(label="Pixel Intensity")
    # ROI_output_filename = f'{image_filename}_background_Intensity.png'
    # output_filepath = os.path.join(save_folder, ROI_output_filename)
    # plt.savefig(output_filepath, dpi=300)
    # plt.close()
    return mean_intensity
"""_____________________________________RBF algorithm__________________________________"""
# Initialization
beta = 1
alpha = 0.1
flag = 0
er = 0.01
b= 3.13468228
lens = 120
w = np.random.rand(1,20)
c=np.random.rand(4,20)
c_norm=np.random.rand(4,20)
num = c.shape[1]
u = 0.01*np.random.rand(1,lens)
x = np.zeros((4,lens))
x_norm = np.zeros((4,lens))
Time = np.zeros((1,lens))
r_in = np.random.rand(1,lens)
r_1 = np.random.rand(1,lens)
y_out = np.random.rand(1,lens)
e = np.random.rand(1,lens)
phi = np.random.rand(lens,num)
A = np.random.rand(1,lens)
y_out = np.random.rand(1,lens)
run_time = 1
coefficient=55
u[0,0]=0
f = 5
period=2*f
vertical=170
amplitude=150
# Number of points
ts = 4.8
PWM_time = 2
n = 0
n_1=0
Ref = []
Er = []
V = []
AP = []
real_time_ROI_inrensity = []
real_time_bacground_inrensity = []
T = []
T1 = []
T2 = []
TT1 = []
DQDT2 = []
calculating_V = []
k_e = []
k_e.append(0)
alpha_list = []
B = []
W = []
C = []
PHI = []
X = []
X_norm = []
C_norm = []
# electrode_number = 65536   #A1
# electrode_number = 4194304   #A2
# electrode_number = 64   #A4
# electrode_number = 4096   #B3
# electrode_number = 512   #C4
# electrode_number = 8192   #C1
electrode_number = 128   #D1
# electrode_number = 2048   #F1
# electrode_number = 1048576   #F3
# electrode_number = 16384   #F4
R = 6
# PLot
fig, axs = plt.subplots(2,2,figsize=(17,10), gridspec_kw={'width_ratios': [1, 1]})
# fig.suptitle('Control when gradient $\partial y / \partial u = 1$')
# ax = fig.add_subplot(2, 2, 4, projection='3d')
# axs1=axs[0,0].twinx()
# axs2=axs[0,1].twinx()
plt.ion()


# CLOSED LOOP CONTROL
# Set the output to the desired voltage
send_ascii_data(ser, f"{PWM_time},{electrode_number},{u[0,0]},0\n")
response = receive_data(ser, bytes_to_read=1000)
time.sleep(PWM_time)
send_ascii_data(ser, "0\n")  #need check these two commond, will the 2nd line wait the voltage finished or run immediately
# response = receive_data(ser, bytes_to_read=1000)
send_ascii_data(ser, "2\n")
# response = receive_data(ser, bytes_to_read=1000)
send_ascii_data(ser, "2\n")
# response = receive_data(ser, bytes_to_read=1000)

image_filename = Zen_capture()
Y_M = get_YM_value(image_filename)
I_bacground = get_Background_value(image_filename)
real_time_ROI_inrensity.append(Y_M)
real_time_bacground_inrensity.append(I_bacground)
y_out[0,0] = (Y_M - I_bacground)/I_bacground*100
# print(y_out[0,0])
x[1,0] = y_out[0,0]
AP.append(y_out[0,0])
V.append(u[0,0])
W.append(w)
C.append(c)
B.append(b)
PHI.append(phi)
X.append(x)
calculating_V.append(u[0,0])
for t in range(lens-1):
    time1=time.time()
    Time[0,t] = t*ts
    T.append(Time[0,t])
    r_in[0,t] = R
    x[0,0]=r_in[0,0]
    r_1[0,t] = R
    x_norm[:,t] = (x[:,t]/np.max(np.abs(x[:,t]))+1e-8)
    c_norm[:,:] = c[:,:]/np.max(np.abs(c[:,:]))
    X_norm.append(x_norm.copy())
    C_norm.append(c_norm.copy())
    e[0,t] = (y_out[0,t] - r_in[0,t])
    Ref.append(r_in[0,t])
    Er.append(e[0,t])

    for i in range(num):
        phi[t,i] = np.exp((-np.sum(np.square(x_norm[:,t]-c_norm[:,i])))/(2*beta**2))
    u[0,t+1] = (np.sum(w*np.transpose(phi[t,:])) + b)*0.5
    # print("phi", phi)
    # print("u[0,t+1]", u[0,t+1])
    calculating_V.append(u[0,t+1])
    if u[0,t+1] >= 0:
        u[0,t+1] = min(u[0,t+1],0)
    elif u[0,t+1] < 0:
        u[0,t+1] = max(u[0,t+1],-2.2)
    else:
        u[0,t+1] = u[0,t+1]

    V.append(u[0,t+1])
    send_ascii_data(ser, f"{PWM_time},{electrode_number},{u[0,t+1]},0\n")
    time.sleep(PWM_time)
    
    send_ascii_data(ser, "0\n")
    # response = receive_data(ser, bytes_to_read=1000)
    send_ascii_data(ser, "2\n")
    # response = receive_data(ser, bytes_to_read=1000)
    send_ascii_data(ser, "2\n")
    # response = receive_data(ser, bytes_to_read=1000)
    time_after_PWM=time.time()
    

    image_filename = Zen_capture()
    Y_M = get_YM_value(image_filename) 
    I_bacground = get_Background_value(image_filename)
    y_out[0,t+1] = (Y_M - I_bacground)/I_bacground*100
    time_after_capture_extract_intensity=time.time()

    
    AP.append(y_out[0,t+1])
    real_time_ROI_inrensity.append(Y_M)
    real_time_bacground_inrensity.append(I_bacground)

    # print(f"Loop t={t}, i={i}")
    # print(f"e[0, t] = {e[0, t]}")
    # print(f"w[0, i] = {w[0, i]}")
    # print(f"w[0, i] = {w[0, i]}")
    # print(f"phi[t, i] = {phi[t, i]}")
    # print(f"Update amount = {alpha * e[0, t] * w[0, i] * (phi[t, i] / (beta * beta)) * (x_norm[:,t]-c_norm[:,i])}")


    for i in range(num):
        delta_c = alpha * e[0, t] * w[0, i] * (phi[t, i] / beta**2) * (x_norm[:,t]-c_norm[:,i])
        max_step = 1.0  # or any value you feel safe with
        delta_c = np.clip(delta_c, -max_step, max_step)
        c[:, i] += delta_c
    w += alpha * e[0, t] * phi[t, :]
    w = np.clip(w, -10, 10)  # or even narrower, like [-5, 5]

    b = b + e[0,t]*alpha
    
    # print(f"c[:, {i}] after update: {c[:, i]}")

    W.append(w.copy())
    C.append(c.copy())
    B.append(b)

    flag = flag + 1

    x[0,t+1] = r_1[0,t]
    x[1,t+1] = y_out[0,t+1]
    x[2,t+1] = x[0,t]
    x[3,t+1] = x[1,t]
    print("y_out[0,t+1] ",y_out[0,t+1] )
    print("x[1,t+1] ",x[1,t+1] )
    PHI.append(phi)
    X.append(x[:, t+1].copy())

    if flag>=7:
        me = np.sum(np.abs(Er[5:]))/len(Er[5:])
    else:
        me = 0
    clear_output(wait = True)
    if flag>=2:
        k_e.append(e[0,t]-e[0,t-1])
        
    else:
        pass
    alpha_list.append(alpha)
    # self.learning_rate *= (1 + self.learning_rate_adjustment_factor * np.abs(error))

    time2=time.time()
    
    # plt.clf()  ##Clear figure
    # axs1.cla()  ## Clear axis
    # plt.xticks(rotation=40)
    plt.cla()
    axs[0,0].clear()
    axs[0,1].clear()
    axs[1,0].clear()
    axs[1,1].clear()

    T1.append(mdates.date2num(datetime.datetime.strptime(str(datetime.datetime.now()),'%Y-%m-%d %H:%M:%S.%f')))
    T2.append(str(datetime.datetime.now().time().hour)+':'+str(datetime.datetime.now().time().minute)+':'+str(datetime.datetime.now().time().second))
    
    axs[0,0].plot(T,AP[0:-1],label="Relative fluorescence intensity:%.4f" % AP[-1],color='orange',linewidth=4,markersize=12,marker='d')
    axs[0,0].plot(T,Ref,label="Reference:%.4f" % Ref[-1],color='black',linewidth=4,markersize=8,marker='o',alpha=0.5)
    axs[0,1].plot(T,real_time_ROI_inrensity[0:-1],label="real_time_ROI_inrensity:%.4f" % Er[-1],color='black',linewidth=4, markersize=10,marker='s',alpha=0.5)
    # axs[0,1].axhline(me,color='black',ls='-.',label="Average Error:%.4f" % me)
    axs[0,1].plot(T,real_time_bacground_inrensity[0:-1],label="real_time_bacground_inrensity:%.4f" % k_e[-1],color='teal',linewidth=4, markersize=10,marker='s',alpha=0.5)
    axs[1,0].plot(T,V[0:-1],label="Voltage:%.4f V" % V[-1],linewidth=4,color='violet',markersize=12,marker='s')
    axs[1,0].plot(T,calculating_V[0:-1],label="Calculating Voltage:%.4f V" % V[-1],linewidth=4,color='violet',markersize=12,marker='s',alpha=0.5)
    axs[1,1].plot(T,alpha_list,label="Learning rate:%.4f " % alpha_list[-1],color='teal',linewidth=4, markersize=10,marker='s',alpha=0.2)

    axs[0,0].set_xlabel('Time (s)', color = '#000000')
    axs[0,0].set_ylabel('Real time Relative fluorescence intensity (%)', color = '#000000')
    # axs1.set_ylabel('Real time Relative fluorescence intensity (%)', color = 'orange', fontweight="bold")
    # axs1.tick_params(labelcolor='orange', labelsize='medium', width=2)
    axs[0,0].ticklabel_format(useOffset=False)
    axs[0,0].grid(True, linestyle='-.')
    axs[0,0].tick_params(labelcolor='#000000', labelsize='medium', width=2)
    # axs[0,0].xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
    axs[0,0].legend(loc = 'lower right', fontsize="10", framealpha=0)
    axs[0,1].set_xlabel('Time (s)', color = '#000000')
    axs[0,1].set_ylabel('Intensity (a.u.)', color = '#000000')
    axs[0,1].ticklabel_format(useOffset=False)
    axs[0,1].grid(True, linestyle='-.')
    axs[0,1].tick_params(labelcolor='#000000', labelsize='medium', width=2)
    axs[0,1].legend(loc = 'upper right', fontsize="10", framealpha=0)
    # axs2.set_ylabel('real_time_bacground_inrensity (a.u.)', color = 'teal', fontweight="bold")
    # axs2.yaxis.set_label_coords(1.1, 0.5)
    # axs2.tick_params(labelcolor='teal', labelsize='medium', width=2)
    axs[1,0].set_xlabel('Time (s)', color = '#000000')
    axs[1,0].set_ylabel('Voltage applied (V)', color = '#000000')
    axs[1,0].ticklabel_format(useOffset=False)
    axs[1,0].grid(True, linestyle='-.')
    axs[1,0].tick_params(labelcolor='#000000', labelsize='medium', width=2)
    axs[1,0].legend(loc = 'lower right', fontsize="10", framealpha=0)
    axs[0,0].spines["top"].set_linewidth(2)    # Top border
    axs[0,0].spines["bottom"].set_linewidth(2) # Bottom border
    axs[0,0].spines["left"].set_linewidth(2)   # Left border
    axs[0,0].spines["right"].set_linewidth(2)  # Right border
    # axs1.spines["right"].set_linewidth(2) 
    # axs1.spines["right"].set_edgecolor("orange")
    # axs1.tick_params(axis='y',      # Apply to both x and y axes
    #            which='major',    # Apply to major ticks
    #            width=2,          # Set tick width
    #            color='orange',      # Set tick color
    #            length=4)        # Set tick length
    axs[0,1].spines["top"].set_linewidth(2)    # Top border
    axs[0,1].spines["bottom"].set_linewidth(2) # Bottom border
    axs[0,1].spines["left"].set_linewidth(2)   # Left border
    axs[0,1].spines["right"].set_linewidth(2)  # Right border  
    # axs2.spines["right"].set_linewidth(2) 
    # axs2.spines["right"].set_edgecolor("teal")
    # axs2.tick_params(axis='y',      # Apply to both x and y axes
    #            which='major',    # Apply to major ticks
    #            width=2,          # Set tick width
    #            color='teal',      # Set tick color
    #            length=4)        # Set tick length  
    axs[1,0].spines["top"].set_linewidth(2)    # Top border
    axs[1,0].spines["bottom"].set_linewidth(2) # Bottom border
    axs[1,0].spines["left"].set_linewidth(2)   # Left border
    axs[1,0].spines["right"].set_linewidth(2)  # Right border

    axs[1,1].spines["top"].set_linewidth(2)    # Top border
    axs[1,1].spines["bottom"].set_linewidth(2) # Bottom border
    axs[1,1].spines["left"].set_linewidth(2)   # Left border
    axs[1,1].spines["right"].set_linewidth(2)  # Right border
    axs[1,1].set_xlabel('Time (s)', color = '#000000')
    axs[1,1].set_ylabel('Learning rate (a.u.)', color = '#000000')
    axs[1,1].ticklabel_format(useOffset=False)
    axs[1,1].grid(True, linestyle='-.')
    axs[1,1].tick_params(labelcolor='#000000', labelsize='medium', width=2)
    axs[1,1].legend(loc = 'lower right', fontsize="10", framealpha=0)
    plt.subplots_adjust(left=0.1,
                    bottom=0.1,
                    right=0.9,
                    top=0.94,
                    wspace=0.2,
                    hspace=0.3)
    
    plt.pause(0.005)
    time3=time.time()
    # time.sleep(4-(time3-time1))
    print("****Time of one loop (read+calculate+plot)****", time3-time1)
    print("****Time of loop to after PWM****", time_after_PWM-time1)
    print("****Time of PWM to capture and extract intensity****", time_after_capture_extract_intensity-time_after_PWM)
    print("****Time of RBF calculating****", time2-time_after_capture_extract_intensity)
    print("****Time of plot****", time3-time2)
filename_1='C:/Users/PingLab_PC8/Desktop/XinZhang/Zen_python_control/1/alpha_%s_experiment_length_%s_PWM_electronumber_number_%s_cycle_period%s_PWM_time%s.png' % (alpha,lens,electrode_number,ts,PWM_time)
root, ext = os.path.splitext(filename_1)
while os.path.exists(filename_1):
    n_1 += 1
    filename_1='%s_%i%s' % (root, n_1, ext)
plt.savefig(filename_1)

send_ascii_data(ser, "0\n")
response = receive_data(ser, bytes_to_read=1000)
"""___________________________Save as csv___________________________"""
fields = ['Time_series', 'T1','Real time', 'Reference', 'Measured ROI Intensity', 'Measured Background Intensity', 'relative intensity y_out[0,t+1]', 'Er e[0,t]',"Real Voltage u[0,t+1]",'calculating_V from RBF u[0,t+1]','weight','center','bias','phi','X','X_norm','C_norm']
filename='C:/Users/PingLab_PC8/Desktop/XinZhang/Zen_python_control/1/alpha_%s_experiment_length_%s_PWM_electronumber_number_%s_cycle_period%s_PWM_time%s.csv' % (alpha,lens,electrode_number,ts,PWM_time)
root, ext = os.path.splitext(filename)
while os.path.exists(filename):
    n += 1
    filename='%s_%i%s' % (root, n, ext)
outfile = open(filename,'w', newline='')
out = csv.writer(outfile)
out.writerow(fields)
out.writerows(zip_longest(*[T, T1, T2, Ref, real_time_ROI_inrensity, real_time_bacground_inrensity, AP, Er,V, calculating_V,W, C, B,PHI,X,X_norm,C_norm]))
outfile.close()
plt.ioff()    
plt.show() 





