using Eplan.EplApi.DataModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPLAN_API.SAP
{
    public class GEC
    {
        public string ID;
        public uint value;
        public String sValue;
        public string name;
        public bool isNumeric = true;
        public bool active = false;

        public GEC(string ID)
        {
            this.ID = ID;
        }
        public GEC(string ID, uint value)
        {
            this.ID = ID;
            this.value = value;
        }
        public GEC(string ID, string name)
        {
            this.ID = ID;
            this.name = name;
            this.value = 0;
        }
        public GEC(string ID, uint value, string name)
        {
            this.ID = ID;
            this.name = name;

            if (value == 99)
                this.value = 0;
            else
                this.value = value;
        }

        public GEC(string ID, String sValue, string name)
        {
            this.ID = ID;
            this.sValue = sValue;
            this.name = name;
        }

        public String getValue()
        {
            if (sValue != null)
                return sValue;
            else
                return value.ToString();
        }
        public String getName()
        {
            return name;
        }
        public String getID()
        {
            return ID;
        }

        public void setValue(uint value)
        {
            this.value = value;
            active = true;
        }

        public void setValue(string value)
        {
            this.sValue = value;
            isNumeric = false;
            active = true;
        }

        public String ValueToText()
        {

            // Funcion para buscar el nombre del parametro GEC en archivo CSV
            string textValue = " ";
            if (isNumeric)
            {
                PathInfo EPLANPaths = new PathInfo();
                //string path = @String.Concat(oProject.DocumentDirectory.Substring(0, oProject.DocumentDirectory.Length - 3), "GEC\\ID_Functions.csv");
                string path = String.Concat(EPLANPaths.Documents, "\\ID_Functions.csv");

                //string fileName = "GEC_parameters.csv";
                string[] filas = File.ReadAllLines(path);

                foreach (var fila in filas)
                {
                    string[] sfila = fila.Split(';');
                    if (sfila[0] != "")
                        if (value.ToString().Equals(sfila[0]))
                        {
                            textValue = sfila[1];
                            break;
                        }
                        else
                            textValue = value.ToString();
                }
            }
            else
                textValue = sValue;

            return textValue;
        }

        public enum Param
        {
            Empty = 0,
            Motor_speed_input_1 = 1,
            Motor_speed_input_2 = 2,
            Main_shaft_speed_monitor_3 = 3,
            Main_shaft_speed_monitor_1 = 4,
            Main_shaft_speed_monitor_2 = 5,
            Up_maint_order = 6,
            Down_maint_order = 7,
            Top_open_floor_plate_1 = 8,
            Top_open_floor_plate_2 = 9,
            Bottom_open_floor_plate_1 = 10,
            Bottom_open_floor_plate_2 = 11,
            Flywheel_protection_1 = 12,
            Flywheel_protection_2 = 13,
            Contactor_FB_3 = 14,
            Contactor_FB_4 = 15,
            Contactor_FB_5 = 16,
            Contactor_FB_6 = 17,
            Contactor_FB_7 = 18,
            Contactor_FB_8 = 19,
            Contactor_FB_brake_M1 = 20,
            Contactor_FB_brake_M2 = 21,
            Contactor_FB_aux_brake = 22,
            Contactor_FB_delay_aux_brake = 24,
            _220V_AC_voltage_monitoring = 25,
            Up_key_order = 26,
            Down_key_order = 27,
            Brake_function_brake_3_mot_1 = 28,
            Brake_function_brake_4_mot_1 = 29,
            Brake_function_brake_1_mot_2 = 30,
            Brake_function_brake_2_mot_2 = 31,
            Brake_function_brake_3_mot_2 = 32,
            Brake_function_brake_4_mot_2 = 33,
            Aux_brake_status_1 = 34,
            Aux_brake_status_2 = 35,
            Drive_chain_DuTriplex = 36,
            Drive_chain_2_DuTriplex = 37,
            Chain_locking_device_SS = 38,
            HR_drive_chain_upper_left = 39,
            HR_drive_chain_upper_right = 40,
            HR_drive_chain_lower_left = 41,
            HR_drive_chain_lower_right = 42,
            Tandem_safety_input_1 = 43,
            Tandem_safety_input_2 = 44,
            Step_chain_lock_device = 45,
            Main_shaft_1_displacement = 46,
            Main_shaft_2_displacement = 47,
            _3rd_nominal_speed_selection_with_warning = 48,
            Safety_curtain = 50,
            Switch_disable_safety_curtain = 51,
            Top_removable_barrier = 52,
            Bottom_removable_barrier = 53,
            _2nd_nominal_speed_selection_with_warning = 54,
            Water_pre_level_warning_bottom = 58,
            Water_pre_level_warning_top = 59,
            Asymmetry_phase_control_relay = 60,
            Protection_switch_drive_M1 = 61,
            Fire_alarm_smoke_detector_1 = 62,
            Water_detection_bottom = 63,
            Oil_level_in_pump_1 = 64,
            Bypass_VFD = 65,
            Trundle_for_soft_stop = 66,
            Overtemperature_M1 = 67,
            Overtemperature_M2 = 68,
            VFD_EEC = 69,
            Protection_switch_drive_M2 = 70,
            Broken_handrail_R = 71,
            Broken_handrail_L = 72,
            Brake_wear_brake_1_M1 = 73,
            Brake_wear_brake_2_M1 = 74,
            Brake_wear_brake_1_M2 = 75,
            Brake_wear_brake_2_M2 = 76,
            Brake_wear_brake_3_M1 = 77,
            Brake_wear_brake_4_M1 = 78,
            Brake_wear_brake_3_M2 = 79,
            Brake_wear_brake_4_M2 = 80,
            Earthquake = 81,
            Extra_fault = 82,
            Floor_plate_anti_thieft_Top = 83,
            Stop_switch_for_safety_curtain = 84,
            Diagnostic_safety_curtain_error = 85,
            Thermostat = 86,
            _2nd_high_speed_selection = 87,
            Automatic_key_top = 88,
            Continuous_key_top = 89,
            Remote_key_top = 90,
            Local_key_top = 91,
            Top_radar_right_NC = 92,
            Top_radar_right_NO_only_safety_radar = 93,
            Top_radar_left_NC = 94,
            Top_radar_left_NO_only_safety_radar = 95,
            Top_light_barrier_NC = 96,
            Top_light_barrier_NO = 97,
            Top_light_barrier_comb_plate_NC = 98,
            Top_light_barrier_comb_plate_NO = 99,
            Bottom_radar_right_NC = 100,
            Bottom_radar_right_NO_only_safety_radar = 101,
            Bottom_radar_left_NC = 102,
            Bottom_radar_left_NO_only_safety_radar = 103,
            Bottom_light_barrier_NC = 104,
            Bottom_light_barrier_NO = 105,
            Bottom_light_barrier_comb_plate_NC = 106,
            Bottom_light_barrier_comb_plate_NO = 107,
            User_stop_non_safety = 108,
            Heating_combs_fault = 109,
            Heating_truss_fault = 110,
            Bidirectional = 111,
            Oil_level_in_gearbox = 112,
            Motor_displacement_M1_1 = 113,
            Motor_displacement_M1_2 = 114,
            Motor_displacement_M2_1 = 115,
            Motor_displacement_M2_2 = 116,
            Gearbox_vibration_1 = 117,
            Gearbox_vibration_2 = 118,
            Testing_switch_aux_brake = 119,
            Reduced_inspection_space_switch_bottom = 120,
            Floor_plate_anti_thieft_Bottom = 121,
            Fire_alarm_smoke_detector_2 = 122,
            Fire_alarm_smoke_detector_3 = 123,
            Fire_alarm_smoke_detector_4 = 124,
            FPGA_fault = 125,
            Automatic_key_bottom = 126,
            Continuous_key_bottom = 127,
            Remote_Key_bottom = 128,
            Local_key_bottom = 129,
            Top_up_key_order = 130,
            Top_down_key_order = 131,
            Bottom_up_key_order = 132,
            Bottom_down_key_order = 133,
            Safety_input_configurable_1 = 134,
            Safety_input_configurable_2 = 135,
            Safety_input_configurable_3 = 136,
            Safety_input_configurable_4 = 137,
            Safety_input_configurable_5 = 138,
            Safety_input_configurable_6 = 139,
            Safety_input_configurable_7 = 140,
            Safety_input_configurable_8 = 141,
            Safety_input_configurable_9 = 142,
            Safety_input_configurable_10 = 143,
            Top_operational_stop_local_B16 = 144,
            Bottom_operational_stop_Local_B16 = 145,
            Top_operational_stop_remote_B17 = 146,
            Bottom_operational_stop_remote_B17 = 147,
            Water_detection_top = 148,
            Unload_guide_rail = 149,
            Top_caterpillar_right_SS = 150,
            Top_caterpillar_left_SS = 151,
            Top_emergency_stop_external_SS = 152,
            Top_emergency_stop_trolley_SS = 153,
            Shutter_rolling_door_SS = 154,
            Top_buggy_right_SS = 155,
            Top_buggy_left_SS = 156,
            Top_vertical_comb_plate_right_SS = 157,
            Top_vertical_comb_plate_left_SS = 158,
            Top_skirt_right_SS = 159,
            Top_skirt_left_SS = 160,
            Top_skirt_inclined_right_SS = 161,
            Top_skirt_inclined_left_SS = 162,
            Intermediate_broken_step_SS = 163,
            Bottom_emergency_stop_external_SS = 164,
            Bottom_emergency_stop_trolley_SS = 165,
            Mechanical_pawl_brake_SS = 166,
            Bottom_buggy_right_SS = 167,
            Bottom_buggy_left_SS = 168,
            Bottom_vertical_comb_plate_right_SS = 169,
            Bottom_vertical_comb_plate_left_SS = 170,
            Bottom_glass_outer_cladding_right_SS = 171,
            Bottom_glass_outer_cladding_left_SS = 172,
            Bottom_skirt_right_SS = 173,
            Bottom_skirt_left_SS = 174,
            Bottom_skirt_inclined_right_SS = 175,
            Bottom_skirt_inclined_left_SS = 176,
            Bottom_shutter_rolling_doors_SS = 177,
            Up_remote = 178,
            Down_remote = 179,
            Stop_remote = 180,
            Reset_remote = 181,
            Linear_heat_detector = 182,
            RGB_sync_input = 183,
            Oil_level_in_pump_2 = 184,
            Reduced_inspection_space_switch_top = 185,
            Continuous_broken_step_SS = 186,
            Continuous_broken_step_CB = 187,
            Oil_level_1_warning = 188,
            Oil_level_2_warning = 189,
            Trolley_top_emergency_stop_2_sec = 191,
            Trolley_bottom_emergency_stop_2_sec = 192,
            Operation_mode_key_1 = 194,
            Order_brake_VFD = 195,
            K11_contactor_feedback = 196,
            EB_end_of_safety_string = 197,
            Speed_selection_output_1 = 200,
            Speed_selection_output_2 = 201,
            Main_2 = 202,
            Fire_alarm = 203,
            Bell = 204,
            Up_indication = 205,
            Down_indication = 206,
            Maintenance_indication = 207,
            Fault_indication = 208,
            Emergency_stop_indication = 209,
            User_stop_indication = 210,
            Lighting_1 = 211,
            Lighting_2 = 212,
            Lighting_3 = 213,
            Reserve = 214,
            Trundle_for_soft_stop_2 = 215,
            Bell_in_control_box_Taiwan_project = 216,
            Safety_Curtain_Control = 217,
            Heating_combs_control = 218,
            Heating_truss_control = 219,
            Heating_handrail_control = 220,
            Heating_motor_control = 221,
            Oil_pump_activation = 222,
            Oil_pump_control_1 = 223,
            Oil_pump_control_2 = 224,
            Stop_indication = 226,
            Cover_open_emergency_stop_top = 227,
            Cover_open_emergency_stop_bottom = 228,
            Cover_open_emergency_stop_intermediate = 229,
            Buzzer_top = 230,
            Buzzer_bottom = 231,
            MAX_output_1 = 232,
            MAX_output_2 = 233,
            MAX_output_3 = 234,
            MAX_output_4 = 235,
            Right_handrail_speed_monitoring_fault = 236,
            Left_handrail_speed_monitoring_fault = 237,
            Floor_plate_anti_thieft_top = 238,
            Floor_plate_anti_thieft_bottom = 239,
            Double_channel_oil_pump_motor_ctrl = 240,
            RGB_sync_output = 248,
            RGB_clock = 250,
            RGB_data = 251,
            Top_traffic_light_green = 252,
            Top_traffic_light_red = 253,
            Bottom_traffic_light_green = 254,
            Bottom_traffic_light_red = 255,
            Top_traffic_light_red_disabled_in_ready_state = 256,
            Bottom_traffic_light_red_disabled_in_ready_state = 257,
            Delta = 260,
            Run_order_VFD = 261,
            Star = 262,
            Main_1 = 263,
            Contactor_brake_M1 = 264,
            Contactor_brake_M2 = 265,
            Delay_contactor_aux_brake = 266,
            Speed_selection_output_3 = 267,
            Pit_water_detection_indication = 268,
            Energy_conservation_mode_indication = 269,
            Out_of_service_Sydney = 270,
            Buzzer_during_start = 271,
            Out_of_service_recovery_fault_Sydney = 272,
            Light_barrier_radar_detection_output_top = 273,
            Light_barrier_radar_detection_output_bottom = 274,
            Oil_pump_pulse_ctrl_1 = 275,
            Oil_pump_pulse_ctrl_2 = 276,
            Esc_InServ = 277,
            Traffic_light_2_direction_mode = 278,
            Red_upper_traffic_light = 279,
            Red_lower_traffic_light = 280,
            Upper_buzzer_including_starting_notification = 281,
            Lower_buzzer_including_starting_notification = 282,
            Emergency_fault_indication = 283,
            Normal_fault_indication = 284,
            General_buzzer = 285,
            Top_PT100 = 300,
            Bottom_PT100 = 301,
            Top_Temperature_Control_4_20mA = 302,
            Top_Traffic_Lights_Supervision_4_20mA_1 = 303,
            Top_Traffic_Lights_Supervision_4_20mA_2 = 304,
            Bottom_Temperature_Control_4_20mA = 305,
            Bottom_Traffic_Lights_Supervision_4_20mA_1 = 306,
            Bottom_Traffic_Lights_Supervision = 307,
            Fixed_comb_plate_removed_top = 308,
            Fixed_comb_plate_removed_bottom = 309,
            Top_horizontal_comb_plate_right = 310,
            Top_horizontal_comb_plate_left = 311,
            Top_handrail_inlet_right = 312,
            Top_handrail_inlet_left = 313,
            Top_pit_emergency_stop = 314,
            Top_broken_step = 315,
            Top_emergency_stop = 316,
            Bottom_horizontal_comb_plate_right = 317,
            Bottom_horizontal_comb_plate_left = 318,
            Bottom_handrail_inlet_right = 319,
            Bottom_handrail_inlet_left = 320,
            Bottom_pit_emergency_stop = 321,
            Bottom_broken_step = 322,
            Bottom_emergency_stop = 323,
            Bottom_step_chain_right = 324,
            Bottom_step_chain_left = 325,
            UP_order = 326,
            Down_order = 327,
            Handrail_left_speed_sensor = 328,
            Handrail_right_speed_sensor = 329,
            Missing_step_lower_sensor = 330,
            Missing_step_upper_sensor = 331,
            Reset_button = 332,
            Begin_of_safety_string = 333,
            Contactor_FB_1 = 334,
            Contactor_FB_2 = 335,
            Main_brake_status_M1_1 = 336,
            Main_brake_status_M1_2 = 337,
            DBU_begin_of_safety_string = 338,
            DBL_begin_of_safety_string = 339,
            DBM1_begin_of_safety_string = 340,
            DBM2_begin_of_safety_string = 341,
            Fan_alarm = 342,
            Control_box_temp_alarm = 343,
            Fan_alarm_2s_warning = 344,
            Heating_system_fault = 345,
            Ground_fault = 347,
            Voltage_drop_failure = 348,
            Harmonic_filter_temperature_protection = 349,
            Key_speed_1 = 350,
            Key_speed_2 = 351,
            BMCS_speed_1 = 352,
            BMCS_speed_2 = 353,
            Missing_step_top_T = 354,
            Missing_step_top_B = 355,
            Missing_step_bottom_T = 356,
            Missing_step_bottom_B = 357,
            Additional_truss_fan_fault = 358,
            Truss_low_temp_top = 359,
            Truss_low_temp_bottom = 360,
            Add_HR_overspeed_limit_inp1 = 361,
            Add_HR_overspeed_limit_inp2 = 362,
            Oil_level_gearbox_extra = 363,
            Operation_mode_key_2 = 364,
            Power_outage_main_switch_contact = 365,
            Power_outage_asymmetry_control_relay = 366,
            Inching_comb_plate_top = 367,
            Inching_comb_plate_bottom = 368,
            Electronic_braking_disabled = 369,
            Cover_plate_upper_extra_1 = 370,
            Cover_plate_upper_extra_2 = 371,
            Cover_plate_lower_extra_1 = 372,
            Cover_plate_lower_extra_2 = 373,
            Right_unloading_rail_abrased = 374,
            Left_unloading_rail_abrased = 375,
            Fault_bit_0_3 = 1501,
            Fault_bit_4_7 = 1502,
            Fault_bit_8_11 = 1503,
            Fault_bit_12_15 = 1504,
            Fault_bit_16_19 = 1505,
            Fault_bit_20_23 = 1506,
            Fault_bit_24_27 = 1507,
            Fault_bit_28_31 = 1508,
            Fault_bit_32_35 = 1509,
            Fault_bit_36_39 = 1510,
            Fault_bit_40_43 = 1511,
            Fault_bit_44_47 = 1512,
            Fault_bit_48_51 = 1513,
            Fault_bit_52_55 = 1514,
            Fault_bit_56_59 = 1515,
            Fault_bit_60_63 = 1516,
            Status_bit_0_3 = 1517,
            Status_bit_4_7 = 1518,
            Status_bit_8_11 = 1519,
            Status_bit_12_15 = 1520,
            Status_bit_16_19 = 1521,
            Status_bit_20_23 = 1522,
            Status_bit_24_27 = 1523,
            Status_bit_28_31 = 1524,
            Warning_bit_0_3 = 1525,
            Warning_bit_4_7 = 1526,
            Warning_bit_8_11 = 1527,
            Warning_bit_12_15 = 1528,
            K11_Contactor = 1529,
            Output_for_inching_speed = 1530,
            Emergency_stop_indication_hold_until_run = 1531,
            Inspection_Norma = 65535
        }

        public enum Code
        {
            NoStandard = 1,
            EN115 = 2,
            ASME_B44 = 3,
            GB_Code = 4
        }

        public enum ContactorFB
        {
            empty = 0,
            K1_1_K1_2 = 3,
            K2_1_K2_2 = 12,
            K2_1_2_K2_2_2 = 192,
            K1_1 = 1,
            K1_2 = 2,
            K2_1_K2_1_1 = 4,
            K2_2_K2_2_1 = 8,
            K10 = 16,
            K10_2_K10_1 = 32,
            K2_1_2 = 64,
            K2_2_2 = 128
        }

        public enum Mode
        {
            No_People_Detection = 0,
            Intermittent = 1,
            Standby = 2,
            Intermittent_Standby = 3,
            TwoDirection = 4
        }

        public enum MotorConnection
        {
            Delta = 0,
            Y_D = 1,
            EEC = 2,
            Smart_inverter_Delta = 3,
            Smart_inverter_Y_D = 4,
            Bypass_inverter_Delta = 5,
            Bypass_inverter_Y_D = 6,
            Nominal_load_inverter_Delta = 7,
            Nominal_load_inverter_Y_D = 8,
            Nominal_load_inverter_no_wired_bypass = 9,
            Bypass_inverter_with_sync_Relay = 10,
            Nominal_load_inverter_Delta_EnergySaving = 11,
            Nominal_load_inverter_Y_D_EnergySaving = 12

        }

        public enum Active
        {
            Disable = 0,
            Enable = 1
        }

        public enum Language
        {
            None = 0,
            English = 1,
            Spanish = 2,
            Chinese = 3,
            German = 4,
            Portuguese = 5,
            French = 6
        }

        public enum Lighting
        {
            off = 0,
            on = 1,
            auto = 2
        }

        public enum MotorSensor
        {
            Two_Sensors_in_motor = 0,
            Korea = 1,
            Two_Sensors_Main_Shaft = 2,
            Two_Sensors_in_motor_Two_Main_Shaft = 3,
            Sensors_in_main_shaft_teeth = 5

        }

        public enum PeopleDetectionSensor
        {
            Two_Input = 0,
            One_Input = 1
        }

        public enum Tandem
        {
            No_Tandem = 0,
            Lower_escalator_tandem = 1,
            Upper_escalator_tandem = 2
        }

    }
}
