import functions as f
import DeviceList as dl
from openpyxl import Workbook

# Scrubbing all data from .csv file
# dict_usable_data = f.read_csv('Source\\AlphaTest.csv')
# dict_usable_data = f.read_csv('Source\\The Grand 091624.csv')
# dict_usable_data = f.read_csv('Source\\The Grand 092324.csv')
dict_usable_data = f.read_csv('Source\\Tempdata.csv')

# for k, v in dict_usable_data.items():
#     print(k)

# Creating new file (workbook)
main_wb = Workbook() 
del main_wb['Sheet'] # Deleting the first sheet that is auto created

# Export Mesh Status data
device_list = dl.GRAND_list_of_all_devices
# device_list = dl.ALPHA_TEST_list_of_all_devices
device_list.sort()

if 'Mesh Status' in dict_usable_data:
    f.Mesh_Status(dict_usable_data, main_wb, device_list, 'Mesh Status')

if 'OutgoingPingRequestNotification' in dict_usable_data:
    f.OutgoingPingRequestNotification(dict_usable_data, main_wb, 'OutgoingPingRequestNotification')

# if 'Mesh' in dict_usable_data:
#     f.export_msg_ctr_sync_response_to_file(dict_usable_data, 'Msg Ctr Sync Response.xlsx')

if 'OutgoingMeshDeviceDiagnosticRequest' in dict_usable_data:
    f.OutgoingMeshDeviceDiagnosticRequest(dict_usable_data, main_wb, 'OutgoingMeshDiagnosticRequest')

if 'Open' in dict_usable_data:
    f.Open(dict_usable_data, main_wb, 'Open')

if 'Close' in dict_usable_data:
    f.Close(dict_usable_data, main_wb, 'Close')

if 'OpenAndClose' in dict_usable_data:
    f.OpenAndClose(dict_usable_data, main_wb, 'OpenAndClose')

if 'OutgoingLockingPlanReadAuditRequest' in dict_usable_data:
    f.OutgoingLockingPlanReadAuditRequest(dict_usable_data, main_wb, 'LockingPlanReadAuditRequest')

if 'OutgoingLockingPlanAuditPointerRequest' in dict_usable_data:
    f.OutgoingLockingPlanAuditPointerRequest(dict_usable_data, main_wb, 'LockingPlanAuditPointerRequest')

if 'OutgoingTraceRtRequest' in dict_usable_data:
    f.OutgoingTraceRtRequest(dict_usable_data, main_wb, 'TraceRtRequest')

if 'OutgoingLockingPlanStateRequest' in dict_usable_data:
    f.OutgoingLockingPlanStateRequest(dict_usable_data, main_wb, 'LockingPlanStateRequest')

if 'OutgoingLockingPlanDSTReadRequest' in dict_usable_data:
    f.OutgoingLockingPlanDSTReadRequest(dict_usable_data, main_wb, 'LockingPlanDSTReadRequest')

if 'OutgoingMeshEventNodeIdsRequest' in dict_usable_data:
    f.OutgoingMeshEventNodeIdsRequest(dict_usable_data, main_wb, 'MeshEventNodeIdsRequest')

if 'LockBlockUnblock' in dict_usable_data:
    f.LockBlockUnblock(dict_usable_data, main_wb, 'LockBlockUnblock')

if 'OutgoingMasterUserCancelRequest' in dict_usable_data:
    f.OutgoingMasterUserCancelRequest(dict_usable_data, main_wb, 'MasterUserCancelRequest')

if 'OutgoingCardErrorDiagnoseRequest' in dict_usable_data:
    f.OutgoingCardErrorDiagnoseRequest(dict_usable_data, main_wb, 'CardErrorDiagnoseRequest')

if 'OutgoingMeshStatusWindowRequest' in dict_usable_data:
    f.OutgoingMeshStatusWindowRequest(dict_usable_data, main_wb, 'MeshStatusWindowRequest')

if 'OutgoingLockingPlanProgrammingRequest' in dict_usable_data:
    f.OutgoingLockingPlanProgrammingRequest(dict_usable_data, main_wb, 'LockingPlanProgrammingRequest')

if 'OutgoingLockingPlanCalendarRequest' in dict_usable_data:
    f.OutgoingLockingPlanCalendarRequest(dict_usable_data, main_wb, 'LockingPlanCalendarRequest')

if 'OutgoingLockingPlanAutomaticChangesRequest' in dict_usable_data:
    f.OutgoingLockingPlanAutomaticChangesRequest(dict_usable_data, main_wb, 'AutomaticChangesRequest')

if 'OutgoingLockingPlanShiftsRequest' in dict_usable_data:
    f.OutgoingLockingPlanShiftsRequest(dict_usable_data, main_wb, 'LockingPlanShiftsRequest')

if 'OutgoingLockingPlanMasterCardKeycodesRequest' in dict_usable_data:
    f.OutgoingLockingPlanMasterCardKeycodesRequest(dict_usable_data, main_wb, 'MasterCardKeycodesRequest')

if 'OutgoingLockingPlanGuestCardKeycodeRequest' in dict_usable_data:
    f.OutgoingLockingPlanGuestCardKeycodeRequest(dict_usable_data, main_wb, 'GuestCardKeycodeRequest')

if 'OutgoingLockingPlanSpecialCardKeycodesRequest' in dict_usable_data:
    f.OutgoingLockingPlanSpecialCardKeycodesRequest(dict_usable_data, main_wb, 'SpecialCardKeycodesRequest')

if 'OutgoingLockingPlanRtcRequest' in dict_usable_data:
    f.OutgoingLockingPlanRtcRequest(dict_usable_data, main_wb, 'LockingPlanRtcRequest')

if 'OutgoingLockingPlanPriorityMsgConfigRequest' in dict_usable_data:
    f.OutgoingLockingPlanPriorityMsgConfigRequest(dict_usable_data, main_wb, 'PriorityMsgConfigRequest')

if 'OutgoingLockingPlanGeneralRequest' in dict_usable_data:
    f.OutgoingLockingPlanGeneralRequest(dict_usable_data, main_wb, 'GeneralRequest')

if 'OutgoingLockDiagnosticGetRequest' in dict_usable_data:
    f.OutgoingLockDiagnosticGetRequest(dict_usable_data, main_wb, 'LockDiagnosticGetRequest')

if 'OutgoingLockingPlanRTCReadRequest' in dict_usable_data:
    f.OutgoingLockingPlanRTCReadRequest(dict_usable_data, main_wb, 'RTCReadRequest')

if 'OutgoingMFGDateRequest' in dict_usable_data:
    f.OutgoingMFGDateRequest(dict_usable_data, main_wb, 'OutgoingMFGDateRequest')

if 'LockingPlanRead' in dict_usable_data:
    f.LockingPlanRead(dict_usable_data, main_wb, 'LockingPlanRead')

if 'ModifyStayDeviceSource' in dict_usable_data:
    f.ModifyStayDeviceSource(dict_usable_data, main_wb, 'ModifyStayDeviceSource')

if 'ModifyStayDeviceDestination' in dict_usable_data:
    f.ModifyStayDeviceDestination(dict_usable_data, main_wb, 'ModifyStayDeviceDestination')

# Sorting the sheets by name
main_wb._sheets.sort(key=lambda ws: ws.title)

# Saving the new workbook and naming it
main_wb.save('Target\\Grand\\The Grand.xlsx') 
# main_wb.save('Target\\Alpha Testing\\main.xlsx')