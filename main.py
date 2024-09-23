import functions as f
import DeviceList as dl
from openpyxl import Workbook

# Scrubbing all data from .csv file
dict_usable_data = f.read_csv('Source\\AlphaTest.csv')
# dict_usable_data = f.read_csv('Source\\The Grand 091624.csv')

# Creating new file (workbook)
main_wb = Workbook() 
del main_wb['Sheet'] # Deleting the first sheet that is auto created

# Export Mesh Status data
# device_list = dl.GRAND_list_of_all_devices
device_list = dl.ALPHA_TEST_list_of_all_devices
device_list.sort()

# f.Mesh_Status(dict_usable_data, main_wb, device_list, 'Mesh Status')

# Export Outgoing Ping Request data
# f.OutgoingPingRequestNotification(dict_usable_data, main_wb, 'OutgoingPingRequestNotification')

# Export Msg Ctr Sync Response data
# f.export_msg_ctr_sync_response_to_file(dict_usable_data, 'Msg Ctr Sync Response.xlsx')

# Export Outgoing Mesh Device Diagnostic Request data
# f.OutgoingMeshDeviceDiagnosticRequest(dict_usable_data, main_wb, 'OutgoingMeshDiagnosticRequest')

# Export Open data
# f.Open(dict_usable_data, main_wb, 'Open')

# Export Close data
# f.Close(dict_usable_data, main_wb, 'Close')

# Export OpenAndClose data
# f.OpenAndClose(dict_usable_data, main_wb, 'OpenAndClose')

# Export LockingPlanReadAuditRequest data
# f.OutgoingLockingPlanReadAuditRequest(dict_usable_data, main_wb, 'LockingPlanReadAuditRequest')

# Export LockingPlanAuditPointerRequest data
# f.OutgoingLockingPlanAuditPointerRequest(dict_usable_data, main_wb, 'LockingPlanAuditPointerRequest')

# Saving the new workbook and naming it
# main_wb.save('Target\\Grand\\The Grand.xlsx')
main_wb.save('Target\\Alpha Testing\\main.xlsx')