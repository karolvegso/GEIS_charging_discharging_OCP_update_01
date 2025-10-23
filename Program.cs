using EcoChemie.Autolab.Extensions.Widgets;
using EcoChemie.Autolab.Sdk;
using EcoChemie.Autolab.Sdk.MultiAutolab;
using EcoChemie.Autolab.Sdk.Nova;
using EcoChemie.Autolab.SDK.Extensions;
using EcoChemie.Shared.MultiAutolabConfiguration;
using EcoChemie.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;
using EI = EcoChemie.Autolab.Sdk.EI;
using Instrument = EcoChemie.Autolab.Sdk.Instrument;

namespace GEIS_charge_discharge_Autolab_ConsoleApp_05
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // print mode information
            Console.WriteLine("Autolab is operating in the galvanostatic mode.");
            // print offset/DC current 
            Console.WriteLine($"The offset current is: {args[0]}");
            float offset_current = float.Parse(args[0]);
            // FRA amplitude current
            Console.WriteLine($"The amplitude current is: {args[1]}");
            float amp_current = float.Parse(args[1]);
            // EIS starting exponent of frequency
            Console.WriteLine($"The starting exponent of frequency is: {args[2]}");
            float freq_exp_start = float.Parse(args[2]);
            // EIS stop exponent of frequency
            Console.WriteLine($"The stop exponent of frequency is: {args[3]}");
            float freq_exp_stop = float.Parse(args[3]);
            // EIS exponent step of frequency
            Console.WriteLine($"The exponent step of frequency is: {args[4]}");
            float freq_exp_step = float.Parse(args[4]);
            // define minimum integration cycles
            Console.WriteLine($"The minimum integration cycles are: {args[5]}");
            int min_integ_cycles = int.Parse(args[5]);
            // define minimum integration time
            Console.WriteLine($"The minimum integration time is: {args[6]}");
            float min_integ_time = float.Parse(args[6]);
            // define maximum voltage threshold for galvanostatic charging
            Console.WriteLine($"The maximum voltage threshold in charging is: {args[7]}");
            float vol_max = float.Parse(args[7]);
            // define minimum voltage threshold in galvanostatic discharging 
            Console.WriteLine($"The minimum voltage threshold in discharging is: {args[8]}");
            float vol_min = float.Parse(args[8]);
            // define maximum number of cycles
            Console.WriteLine($"The maximum number of cycles is: {args[9]}");
            int max_no_cycles = int.Parse(args[9]);
            // define waiting time after setting DC current
            Console.WriteLine($"The waiting time after setting DC current is: {args[10]}");
            int wait_time_DC_current = int.Parse(args[10]);
            // define waiting time for cell on
            Console.WriteLine($"The waiting time after cel on is: {args[11]}");
            int wait_time_cell_on = int.Parse(args[11]);
            // define waiting time for fra on
            Console.WriteLine($"The waiting time after fra on is: {args[12]}");
            int wait_time_fra_on = int.Parse(args[12]);
            // define waiting time at the beginning of EIS measurement
            Console.WriteLine($"The waiting time at the beginning of EIS measurement is: {args[13]}");
            int wait_time_begin_EIS_meas = int.Parse(args[13]);
            // define waiting time after frequency change in EIS measurement
            Console.WriteLine($"The waiting time after frequency change in EIS measurement is: {args[14]}");
            int wait_time_freq_change = int.Parse(args[14]);
            // define waiting time after stop detection
            Console.WriteLine($"The waiting time after stop detection is: {args[15]}");
            int wait_time_after_stop = int.Parse(args[15]);
            // tell if list of frequencies is used
            Console.WriteLine($"The list of frequencies is used Yes/No: {args[16]}");
            string read_freq_list = args[16];
            // tell if list of frequencies is ascending or descending
            Console.WriteLine($"The list of frequencies has direction: {args[17]}");
            string direction_freq_list = args[17];
            // tell electochemical cell current range
            Console.WriteLine($"The electrochemical cell current range is: {args[18]}");
            string current_range = args[18];

            // save voltage data
            // path to voltage data
            string vol_filename = @"c:\GEIS_data\voltage_data.txt";

            // save time data
            // path to time data
            string time_filename = @"c:\GEIS_data\time_data.txt";

            // save charge/discharge switch time data
            string time_switch = @"C:\GEIS_data\time_charge_discharge.txt";

            // path to list of ascending frequencies 
            string path_to_ascending_freq_list = @"c:\GEIS_data\list_of_frequencies_ascending.txt";
            // path to list of descending frequencies 
            string path_to_descending_freq_list = @"c:\GEIS_data\list_of_frequencies_descending.txt";

            // initialize number of frequency points variable, it will be filled later
            int no_freq_points = 0;

            // initialize empty frequency list array
            double[] freq_array = new double[0];

            // decide how to generate frequency data
            // decide if calculated frequencies should be used or list of frequencies
            if (read_freq_list == "yes")
            {
                if (direction_freq_list == "descending")
                {
                    // inialize frequency list
                    List<float> freq_list = new List<float>();
                    // read text file
                    using (StreamReader freq_reader = new StreamReader(path_to_descending_freq_list))
                    {
                        string line;
                        while ((line = freq_reader.ReadLine()) != null)
                        {
                            // read frequency value from text file
                            float freq_value = float.Parse(line);
                            // add frequency value to list
                            freq_list.Add(freq_value);
                        }
                    }
                    // get length of list of frequencies
                    no_freq_points = freq_list.Count;
                    // initialize frequency list array
                    freq_array = new double[no_freq_points];
                    // fill frequency list array
                    for (int index_0 = 0; index_0 < no_freq_points; index_0++)
                    {
                        freq_array[index_0] = Math.Round(freq_list[index_0], 6);
                        Console.WriteLine(freq_array[index_0]);
                    }
                }
                else if (direction_freq_list == "ascending")
                {
                    // inialize frequency list
                    List<float> freq_list = new List<float>();
                    // read text file
                    using (StreamReader freq_reader = new StreamReader(path_to_ascending_freq_list))
                    {
                        string line;
                        while ((line = freq_reader.ReadLine()) != null)
                        {
                            // read frequency value from text file
                            float freq_value = float.Parse(line);
                            // add frequency value to list
                            freq_list.Add(freq_value);
                        }
                    }
                    // get length of list of frequencies
                    no_freq_points = freq_list.Count;
                    // initialize frequency list array
                    freq_array = new double[no_freq_points];
                    // fill frequency list array
                    for (int index_0 = 0; index_0 < no_freq_points; index_0++)
                    {
                        freq_array[index_0] = Math.Round(freq_list[index_0], 6); ;
                        Console.WriteLine(freq_array[index_0]);
                    }
                }
                else
                {
                    // inialize frequency list
                    List<float> freq_list = new List<float>();
                    // read text file
                    using (StreamReader freq_reader = new StreamReader(path_to_descending_freq_list))
                    {
                        string line;
                        while ((line = freq_reader.ReadLine()) != null)
                        {
                            // read frequency value from text file
                            float freq_value = float.Parse(line);
                            // add frequency value to list
                            freq_list.Add(freq_value);
                        }
                    }
                    // get length of list of frequencies
                    no_freq_points = freq_list.Count;
                    // initialize frequency list array
                    freq_array = new double[no_freq_points];
                    // fill frequency list array
                    for (int index_0 = 0; index_0 < no_freq_points; index_0++)
                    {
                        freq_array[index_0] = Math.Round(freq_list[index_0], 6); ;
                        Console.WriteLine(freq_array[index_0]);
                    }
                }
            }
            else if (read_freq_list == "no")
            {
                // calculate number of frequency points
                no_freq_points = (int)Math.Abs((freq_exp_stop - freq_exp_start) / freq_exp_step) + 1;

                // initialize frequency exponent array
                float[] freq_exp_array = new float[no_freq_points];
                // initialize frequency array
                freq_array = new double[no_freq_points];
                // fill frequency array
                for (int index_0 = 0; index_0 < no_freq_points; index_0++)
                {
                    freq_exp_array[index_0] = freq_exp_start + index_0 * freq_exp_step;
                    double freq_value = Math.Pow(10, freq_exp_array[index_0]);
                    freq_value = Math.Round(freq_value, 6);
                    freq_array[index_0] = freq_value;
                    Console.WriteLine(freq_array[index_0]);
                }
            }
            else
            {
                // inialize frequency list
                List<float> freq_list = new List<float>();
                // read text file
                using (StreamReader freq_reader = new StreamReader(path_to_descending_freq_list))
                {
                    string line;
                    while ((line = freq_reader.ReadLine()) != null)
                    {
                        // read frequency value from text file
                        float freq_value = float.Parse(line);
                        // add frequency value to list
                        freq_list.Add(freq_value);
                    }
                }
                // get length of list of frequencies
                no_freq_points = freq_list.Count;
                // initialize frequency list array
                freq_array = new double[no_freq_points];
                // fill frequency list array
                for (int index_0 = 0; index_0 < no_freq_points; index_0++)
                {
                    freq_array[index_0] = Math.Round(freq_list[index_0], 6); ;
                    Console.WriteLine(freq_array[index_0]);
                }
            }

            // initialize EIS array
            double[,] eis_array = new double[no_freq_points, 5];

            // connect to first Autolab instrument
            InstrumentConnectionManager autolab_manager = new InstrumentConnectionManager(@"C:\Program Files\Metrohm Autolab\Autolab SDK 2.1\Hardware Setup Files");
            var my_instrument = autolab_manager.ConnectToFirstInstrument();
            Console.WriteLine("The program is now connected to the first Autolab instrument.");

            // query if my instrument has FRA module
            bool has_fra_module = my_instrument.HasFraModule();
            // print response on the screen
            if (has_fra_module == true)
            {
                Console.WriteLine("The Autolab has FRA module.");
            }
            else
            {
                Console.WriteLine("Autolab has no FRA module.");
            }

            // Set the instrument to be used by the kernel
            my_instrument.SetAsDefault();
            // print serial number of Autolab with FRA module
            Console.WriteLine($"Fra sdk-measurement on {my_instrument.GetSerialNumber()}");

            // initialize counter for charging and discharging events
            int counter_events = 0;
            // set maximum charging and discharging events
            int max_events = 5;

            // switch off eletrochemical cell of Autolab
            my_instrument.Ei.CellOnOff = EI.EICellOnOff.Off;
            // read current voltage or potential
            // default ocp procedure
            const string procedure_ocp = @"C:\GEIS_data\Measure_OCP_Karol.nox";
            // condition for charging or discharging
            string direction = string.Empty;
            // load ocp procedure
            var procedureFilename = my_instrument.LoadProcedure(procedure_ocp);
            var procedureRun = my_instrument.PrepareProcedure(procedure_ocp);
            // do ocp measurement
            //Console.WriteLine("Measuring ocp");
            procedureRun.Measure();
            // Monitor ocp measurement
            do
            {
                //Console.Write(".");
                Task.Delay(1000);
            } while (procedureRun.IsMeasuring);
            //Console.WriteLine("Finished ocp measurement.");
            // Show results
            double vol_current = DumpTables(procedureRun);
            // print initial potential on the battery cell
            Console.WriteLine($"The initial potential on the battery cell is: {vol_current}");
            // wait 1 s
            System.Threading.Thread.Sleep(1000);

            if (vol_current < vol_max)
            {
                // increnment counter for charging and discharging events
                counter_events++;
                Console.WriteLine("Do battery charging.");
                direction = "charging";
                // set offset/DC curent
                my_instrument.Ei.Setpoint = offset_current;
            }
            else
            {
                // increnment counter for charging and discharging events
                counter_events++;
                Console.WriteLine("Do battery discharging.");
                direction = "discharging";
                // set offset/DC curent
                my_instrument.Ei.Setpoint = (-1.0f) * offset_current;
            }

            // define first while loop to check battery cell potential
            // get right direction, charging or discharging
            while (counter_events < max_events)
            {
                // load ocp procedure
                procedureFilename = my_instrument.LoadProcedure(procedure_ocp);
                procedureRun = my_instrument.PrepareProcedure(procedure_ocp);
                // do ocp measurement
                //Console.WriteLine("Measuring ocp");
                procedureRun.Measure();
                // Monitor ocp measurement
                do
                {
                    //Console.Write(".");
                    Task.Delay(1000);
                } while (procedureRun.IsMeasuring);
                //Console.WriteLine("Finished ocp measurement.");
                // Show results
                vol_current = DumpTables(procedureRun);
                // print initial potential on the battery cell
                Console.WriteLine($"The initial potential on the battery cell is: {vol_current}");
                // wait 1 s
                System.Threading.Thread.Sleep(1000);

                if (direction == "charging" && vol_current < vol_max)
                {
                    // increment counter for charging and discharging events
                    counter_events++;
                    Console.WriteLine("Do battery charging.");
                    direction = "charging";
                }
                else if (direction == "charging" && vol_current >= vol_max)
                {
                    // empty counter for charging and discharging events
                    counter_events = 0;
                    Console.WriteLine("Do battery discharging.");
                    direction = "discharging";
                }
                else if (direction == "discharging" && vol_current >= vol_max)
                {
                    // increment counter for charging and discharging events
                    counter_events++;
                    Console.WriteLine("Do battery discharging.");
                    direction = "discharging";
                }
                else if (direction == "discharging" && vol_current < vol_max)
                {
                    // empty counter for charging and discharging events
                    counter_events = 0;
                    Console.WriteLine("Do battery charging.");
                    direction = "charging";
                }
            }

            // set battery cell DC current
            if (direction == "charging")
            {
                // set offset/DC curent
                my_instrument.Ei.Setpoint = offset_current;
            }
            else if (direction == "discharging")
            {
                // set offset/DC curent
                my_instrument.Ei.Setpoint = (-1.0f) * offset_current;
            }
            else
            {
                // set direction
                direction = "unknown";
            }

            // set mode to galvanostatic mode for GEIS
            my_instrument.Ei.Mode = EI.EIMode.Galvanostatic;

            // Select the appropriate CurrentRange
            if (current_range == "CR00_1000A")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR00_1000A;
            }
            else if (current_range == "CR01_100A")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR01_100A;
            }
            else if (current_range == "CR02_80A")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR02_80A;
            }
            else if (current_range == "CR03_50A")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR03_50A;
            }
            else if (current_range == "CR04_40A")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR04_40A;
            }
            else if (current_range == "CR05_20A")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR05_20A;
            }
            else if (current_range == "CR06_10A")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR06_10A;
            }
            else if (current_range == "CR07_1A")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR07_1A;
            }
            else if (current_range == "CR08_100mA")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR08_100mA;
            }
            else if (current_range == "CR09_10mA")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR09_10mA;
            }
            else if (current_range == "CR10_1mA")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR10_1mA;
            }
            else if (current_range == "CR11_100uA")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR11_100uA;
            }
            else if (current_range == "CR12_10uA")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR12_10uA;
            }
            else if (current_range == "CR13_1uA")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR13_1uA;
            }
            else if (current_range == "CR14_100nA")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR14_100nA;
            }
            else if (current_range == "CR15_10nA")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR15_10nA;
            }
            else if (current_range == "CR16_1nA")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR16_1nA;
            }
            else if (current_range == "CR17_100pA")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR17_100pA;
            }
            else if (current_range == "CR18_10pA")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR18_10pA;
            }
            else if (current_range == "CR19_1pA")
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR19_1pA;
            }
            else
            {
                my_instrument.Ei.CurrentRange = EI.EICurrentRange.CR10_1mA;
            }

            // set high stability measurement
            my_instrument.Ei.Bandwidth = EI.EIBandwidth.High_Stability;
            // switch on the Fra-DSG relay on the PGSTAT
            my_instrument.Ei.EnableDsgInput = true;

            // set FRA parameters
            my_instrument.Fra.Amplitude = amp_current;
            my_instrument.Fra.WaveType = FraWaveType.Sine;
            my_instrument.Fra.MinimumIntegrationCycles = min_integ_cycles;
            my_instrument.Fra.MinimumIntegrationTime = min_integ_time;

            // wait 1 s
            System.Threading.Thread.Sleep(1000);

            // empty counter for charging and discharging events
            counter_events = 0;
            // initialize counter
            int counter = 0;
            // main while loop
            while ((direction == "charging" || direction == "discharging") && (counter < 2 * max_no_cycles))
            {
                if (direction == "charging")
                {
                    // set offset/DC curent
                    my_instrument.Ei.Setpoint = offset_current;
                    // wait time after setting DC current
                    System.Threading.Thread.Sleep(wait_time_DC_current);

                    //Console.WriteLine("Measuring GEIS now...");

                    for (int index_0 = 0; index_0 < no_freq_points; index_0++)
                    {
                        if (index_0 == 0)
                        {
                            // switch off eletrochemical cell of Autolab
                            my_instrument.Ei.CellOnOff = EI.EICellOnOff.Off;
                            // read current voltage or potential
                            // load ocp procedure
                            procedureFilename = my_instrument.LoadProcedure(procedure_ocp);
                            procedureRun = my_instrument.PrepareProcedure(procedure_ocp);
                            // do ocp measurement
                            //Console.WriteLine("Measuring ocp");
                            procedureRun.Measure();
                            // Monitor ocp measurement
                            do
                            {
                                //Console.Write(".");
                                Task.Delay(1000);
                            } while (procedureRun.IsMeasuring);
                            //Console.WriteLine("Finished ocp measurement.");
                            // Show results
                            vol_current = DumpTables(procedureRun);
                            // print the potential at the beginning of EIS
                            Console.WriteLine($"The potential at the beginning of EIS is: {vol_current}");
                            // convert current volatge to string
                            string vol_current_str = vol_current.ToString();
                            // wait 1 s
                            System.Threading.Thread.Sleep(1000);

                            // switch on eletrochemical cell of Autolab
                            my_instrument.Ei.CellOnOff = EI.EICellOnOff.On;
                            // wait time to stabilize after switching the cell on
                            System.Threading.Thread.Sleep(wait_time_cell_on);

                            // Switch on the FRA
                            my_instrument.SwitchFraOn();
                            // wait time to stabilize after switching the fra on
                            System.Threading.Thread.Sleep(wait_time_fra_on);

                            // condition
                            if (vol_current < vol_max)
                            {
                                // write current voltage to text file
                                using (StreamWriter volw = System.IO.File.AppendText(vol_filename))
                                {
                                    volw.WriteLine(vol_current_str);
                                }
                                // write timestamp to text file
                                using (StreamWriter timew = System.IO.File.AppendText(time_filename))
                                {
                                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                    timew.WriteLine(timestamp);
                                }
                                // update charging / discharging direction
                                direction = "charging";
                                // set FRA frequency
                                my_instrument.Fra.Frequency = freq_array[index_0];
                                // wait time at the beginning of EIS measurement
                                System.Threading.Thread.Sleep(wait_time_begin_EIS_meas);
                                // perform measurement by starting FRA
                                my_instrument.Fra.Start();
                                // print end of FRA measurement
                                //Console.WriteLine("Measurement finished.");
                                // measure EIS, impedance
                                double Z_freq = my_instrument.Fra.Frequency;
                                double Z_total = my_instrument.Fra.Modulus[0];
                                double Z_phase = my_instrument.Fra.Phase[0];
                                double Z_real = my_instrument.Fra.Real[0];
                                double Z_imag = my_instrument.Fra.Imaginary[0];
                                //double Z_time = my_instrument.Fra.TimeData[my_instrument.Fra.TimeData.Length - 1];
                                // fill EIS array
                                eis_array[index_0, 0] = Z_freq;
                                eis_array[index_0, 1] = Z_real;
                                eis_array[index_0, 2] = Z_imag;
                                eis_array[index_0, 3] = Z_total;
                                eis_array[index_0, 4] = Z_phase;
                                //eis_array[index_0, 5] = Z_time;
                                // increment counter
                                counter += 1;
                                // empty counter for charging and discharging events
                                counter_events = 0;
                            }
                            else
                            {
                                // increment counter for charging and discharging events 
                                counter_events++;
                                if (counter_events >= max_events)
                                {
                                    // update charging / discharging direction
                                    direction = "discharging";
                                    // write timestamp to text file
                                    using (StreamWriter timeswitch = System.IO.File.AppendText(time_switch))
                                    {
                                        string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        timeswitch.WriteLine(timestamp);
                                    }
                                    // empty counter for charging and discharging events
                                    counter_events = 0;
                                    // break for cycle
                                    break;
                                }
                                else if (counter_events < max_events && counter_events >= 0)
                                {
                                    // update charging / discharging direction
                                    direction = "charging";
                                    // break for cycle
                                    break;
                                }
                                else
                                {
                                    // update charging / discharging direction
                                    direction = "charging";
                                    // break for cycle
                                    break;
                                }
                            }
                        }
                        else
                        {
                            // set FRA frequency
                            my_instrument.Fra.Frequency = freq_array[index_0];
                            // wait time after frequency change
                            System.Threading.Thread.Sleep(wait_time_freq_change);
                            // perform measurement by starting FRA
                            my_instrument.Fra.Start();
                            // print end of FRA measurement
                            //Console.WriteLine("Measurement finished.");
                            // measure EIS, impedance
                            double Z_freq = my_instrument.Fra.Frequency;
                            double Z_total = my_instrument.Fra.Modulus[0];
                            double Z_phase = my_instrument.Fra.Phase[0];
                            double Z_real = my_instrument.Fra.Real[0];
                            double Z_imag = my_instrument.Fra.Imaginary[0];
                            //double Z_time = my_instrument.Fra.TimeData[my_instrument.Fra.TimeData.Length - 1];
                            // fill EIS array
                            eis_array[index_0, 0] = Z_freq;
                            eis_array[index_0, 1] = Z_real;
                            eis_array[index_0, 2] = Z_imag;
                            eis_array[index_0, 3] = Z_total;
                            eis_array[index_0, 4] = Z_phase;
                            //eis_array[index_0, 5] = Z_time;

                            // condition to save data
                            if (index_0 == (no_freq_points - 1))
                            {
                                // save EIS data
                                // root folder
                                string root_folder = @"c:\GEIS_data";
                                // filename
                                string filename = "GEIS_scan";
                                // convert counter to string
                                string counter_str = counter.ToString().PadLeft(6, '0');
                                // build full file name
                                string filename_full = filename + "_" + counter_str + ".txt";
                                // build full path to file
                                string full_path = Path.Combine(root_folder, filename_full);

                                // create object StreamWriter
                                using (StreamWriter sw = new StreamWriter(full_path))
                                {
                                    // main for loop to read pixel values
                                    for (int index_1 = 0; index_1 < no_freq_points; index_1++)
                                    {
                                        for (int index_2 = 0; index_2 < 5; index_2++)
                                        {
                                            sw.Write(eis_array[index_1, index_2] + "\t");
                                        }
                                        sw.WriteLine();
                                    }
                                }
                                // switch off eletrochemical cell of Autolab
                                my_instrument.Ei.CellOnOff = EI.EICellOnOff.Off;
                                // Switch off the FRA, good habit to turn things off
                                my_instrument.SwitchFraOff();
                            }
                        }
                    }
                }
                else if (direction == "discharging")
                {
                    // set offset/DC curent
                    my_instrument.Ei.Setpoint = (-1.0f) * offset_current;
                    // wait time after setting DC current
                    System.Threading.Thread.Sleep(wait_time_DC_current);

                    for (int index_0 = 0; index_0 < no_freq_points; index_0++)
                    {
                        if (index_0 == 0)
                        {

                            // switch off eletrochemical cell of Autolab
                            my_instrument.Ei.CellOnOff = EI.EICellOnOff.Off;
                            // read current voltage or potential
                            // load ocp procedure
                            procedureFilename = my_instrument.LoadProcedure(procedure_ocp);
                            procedureRun = my_instrument.PrepareProcedure(procedure_ocp);
                            // do ocp measurement
                            //Console.WriteLine("Measuring ocp");
                            procedureRun.Measure();
                            // Monitor ocp measurement
                            do
                            {
                                //Console.Write(".");
                                Task.Delay(1000);
                            } while (procedureRun.IsMeasuring);
                            //Console.WriteLine("Finished ocp measurement.");
                            // Show results
                            vol_current = DumpTables(procedureRun);
                            // print the potential at the beginning of EIS
                            Console.WriteLine($"The potential at the beginning of EIS is: {vol_current}");
                            // convert current volatge to string
                            string vol_current_str = vol_current.ToString();
                            // wait 1 s
                            System.Threading.Thread.Sleep(1000);

                            // switch on eletrochemical cell of Autolab
                            my_instrument.Ei.CellOnOff = EI.EICellOnOff.On;
                            // wait time to stabilize after switching the cell on
                            System.Threading.Thread.Sleep(wait_time_cell_on);

                            // Switch on the FRA
                            my_instrument.SwitchFraOn();
                            //wait time to stabilize after switching the fra on
                            System.Threading.Thread.Sleep(wait_time_fra_on);

                            // condition
                            if (vol_current > vol_min)
                            {
                                // write current voltage to text file
                                using (StreamWriter volw = System.IO.File.AppendText(vol_filename))
                                {
                                    volw.WriteLine(vol_current_str);
                                }
                                // write timestamp to text file
                                using (StreamWriter timew = System.IO.File.AppendText(time_filename))
                                {
                                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                    timew.WriteLine(timestamp);
                                }
                                // update charging / discharging direction
                                direction = "discharging";
                                // set FRA frequency
                                my_instrument.Fra.Frequency = freq_array[index_0];
                                // wait time at the beginning of EIS measurement
                                System.Threading.Thread.Sleep(wait_time_begin_EIS_meas);
                                // perform measurement by starting FRA
                                my_instrument.Fra.Start();
                                // print end of FRA measurement
                                //Console.WriteLine("Measurement finished.");
                                // measure EIS, impedance
                                double Z_freq = my_instrument.Fra.Frequency;
                                double Z_total = my_instrument.Fra.Modulus[0];
                                double Z_phase = my_instrument.Fra.Phase[0];
                                double Z_real = my_instrument.Fra.Real[0];
                                double Z_imag = my_instrument.Fra.Imaginary[0];
                                //double Z_time = my_instrument.Fra.TimeData[my_instrument.Fra.TimeData.Length - 1];
                                // fill EIS array
                                eis_array[index_0, 0] = Z_freq;
                                eis_array[index_0, 1] = Z_real;
                                eis_array[index_0, 2] = Z_imag;
                                eis_array[index_0, 3] = Z_total;
                                eis_array[index_0, 4] = Z_phase;
                                //eis_array[index_0, 5] = Z_time;
                                // increment counter
                                counter += 1;
                                // empty counter for charging and discharging events
                                counter_events = 0;
                            }
                            else
                            {
                                // increment counter for charging and discharging events 
                                counter_events++;
                                if (counter_events >= max_events)
                                {
                                    // update charging / discharging direction
                                    direction = "charging";
                                    // write timestamp to text file
                                    using (StreamWriter timeswitch = System.IO.File.AppendText(time_switch))
                                    {
                                        string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        timeswitch.WriteLine(timestamp);
                                    }
                                    // empty counter for charging and discharging events
                                    counter_events = 0;
                                    // break for cycle
                                    break;
                                }
                                else if (counter_events < max_events && counter_events >= 0)
                                {
                                    // update charging / discharging direction
                                    direction = "discharging";
                                    // break for cycle
                                    break;
                                }
                                else
                                {
                                    // update charging / discharging direction
                                    direction = "discharging";
                                    // break for cycle
                                    break;
                                }
                            }
                        }
                        else
                        {
                            // set FRA frequency
                            my_instrument.Fra.Frequency = freq_array[index_0];
                            // wait time after frequency change
                            System.Threading.Thread.Sleep(wait_time_freq_change);
                            // perform measurement by starting FRA
                            my_instrument.Fra.Start();
                            // print end of FRA measurement
                            //Console.WriteLine("Measurement finished.");
                            // measure EIS, impedance
                            double Z_freq = my_instrument.Fra.Frequency;
                            double Z_total = my_instrument.Fra.Modulus[0];
                            double Z_phase = my_instrument.Fra.Phase[0];
                            double Z_real = my_instrument.Fra.Real[0];
                            double Z_imag = my_instrument.Fra.Imaginary[0];
                            //double Z_time = my_instrument.Fra.TimeData[my_instrument.Fra.TimeData.Length - 1];
                            // fill EIS array
                            eis_array[index_0, 0] = Z_freq;
                            eis_array[index_0, 1] = Z_real;
                            eis_array[index_0, 2] = Z_imag;
                            eis_array[index_0, 3] = Z_total;
                            eis_array[index_0, 4] = Z_phase;
                            //eis_array[index_0, 5] = Z_time;

                            // condition to save data
                            if (index_0 == (no_freq_points - 1))
                            {
                                // save EIS data
                                // root folder
                                string root_folder = @"c:\GEIS_data";
                                // filename
                                string filename = "GEIS_scan";
                                // convert counter to string
                                string counter_str = counter.ToString().PadLeft(6, '0');
                                // build full file name
                                string filename_full = filename + "_" + counter_str + ".txt";
                                // build full path to file
                                string full_path = Path.Combine(root_folder, filename_full);

                                // create object StreamWriter
                                using (StreamWriter sw = new StreamWriter(full_path))
                                {
                                    // main for loop to read pixel values
                                    for (int index_1 = 0; index_1 < no_freq_points; index_1++)
                                    {
                                        for (int index_2 = 0; index_2 < 5; index_2++)
                                        {
                                            sw.Write(eis_array[index_1, index_2] + "\t");
                                        }
                                        sw.WriteLine();
                                    }
                                }
                                // switch off eletrochemical cell of Autolab
                                my_instrument.Ei.CellOnOff = EI.EICellOnOff.Off;
                                // Switch off the FRA, good habit to turn things off
                                my_instrument.SwitchFraOff();
                            }
                        }
                    }
                }
                else
                {
                    // break while loop
                    break;
                }

                // stop GEIS measurement script
                // cretae path to stop text file
                string path_to_stop_file = @"c:\GEIS_data\stop.txt";
                // open with shared read/write access
                FileStream fs = new FileStream(path_to_stop_file, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
                // create stream reader object
                using (StreamReader sr = new StreamReader(fs))
                {
                    // read first line
                    string firstLine = sr.ReadLine();
                    //Console.WriteLine(firstLine);
                    // condition
                    if (firstLine == "stop")
                    {
                        Console.WriteLine("The main while loop of GEIS measurement is broken.");
                        Console.WriteLine("The GEIS script is stopped");
                        System.Threading.Thread.Sleep(wait_time_after_stop);
                        // break main while loop
                        break;
                    }
                    else
                    {
                        // continue main while loop
                        continue;
                    }
                }
            }

            // switch off eletrochemical cell of Autolab
            my_instrument.Ei.CellOnOff = EI.EICellOnOff.Off;
            // switch off the FRA relay on the PGSTAT
            my_instrument.Ei.EnableDsgInput = false;

            // Switch off the FRA, good habit to turn things off
            my_instrument.SwitchFraOff();

            // Dispose Autolab manager
            autolab_manager.Dispose();
        }

        #region ToTable
        private static IEnumerable<Table> ToTable(IProcedure result)
        {
            foreach (var c in result.Commands.Where(c => c.Signals.Any()))
                yield return ToTable(c);
        }

        private static Table ToTable(ICommand command)
        {

            // All our signals are columns in the table
            var table = new Table();
            var lSignal = new List<IList>();
            var count = 0;

            var signals = command.Signals
                .Where(cp => cp.ValueAsObject is IList)
                .OrderBy(sig => GetOrder(sig.GetName()));

            int GetOrder(string name)
            {
                if (name == "Index")
                    return 10;
                if (name == "CalcTime")
                    return 20;
                if (name == "Correctedtime")
                    return 30;
                if (name == "EI_0.CalcPotential")
                    return 40;
                if (name == "EI_0.CalcCurrent")
                    return 50;
                return 3000;
            }


            foreach (var cp in signals)
            {
                if (cp.ValueAsObject is IList v)
                {
                    table.AddColumn(cp.GetName());  // Add the column for this signal
                    lSignal.Add(cp.ValueAsObject as IList);
                    count = Math.Max(count, v.Count);
                }
            }

            for (int index = 0; index < count; index++)
            {
                var r = lSignal.Select(l =>
                {
                    if (index < l.Count)
                        return l[index].ToString();
                    return "-";
                }).ToArray();
                table.AddRow(r);
            }

            return table;

        }

        private static Table ToTimeTable(IProcedure procedureRun)
        {
            var table = new Table();
            table.AddColumn("Command", "EmbTimeStamp", "TimeStamp");
            table.AddRow(procedureRun.GetName(), "-", procedureRun.GetStartTime().ToString("hh:mm:ss.fff"));

            procedureRun.Commands.ToList().ForEach(c =>
            {
                table.AddRow(c.GetName(), c.GetTimeStampEmbedded().ToString(), c.GetStartTime().ToString("hh:mm:ss.fff"));
            });

            return table;

        }

        private static double DumpTables(IProcedure procedure)
        {
            // read all tables
            var tables = ToTable(procedure);
            // initialize table counter
            int counter_table = 0;
            // get index of potentials
            int potentialIndex = 3;
            // sum of potentials
            double sum_potential = 0;
            // average potential
            double average_potential = 0;
            // initialize counter
            foreach (var table in tables)
            {
                if (counter_table == 1)
                {
                    //Console.WriteLine(table.ToString());
                    //Console.WriteLine(table.ToCSV());
                    string table_csv = table.ToCSV();
                    // Split into lines (rows)
                    string[] lines = table_csv.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                    int counter_line = 0;
                    // Loop through each row
                    foreach (string line in lines)
                    {
                        if (counter_line > 0)
                        {
                            string[] values = line.Split(',');
                            int values_length = values.Length;
                            //Console.WriteLine("Column, EI_0.CalcPotential: ");
                            if (values_length > potentialIndex)
                            {
                                //Console.WriteLine($"Current potentials is: {values[potentialIndex]}");
                                sum_potential += double.Parse(values[potentialIndex]);
                            }
                        }
                        // increment counter potential
                        counter_line++;
                    }
                    // 
                    int counter_potential = counter_line - 1;
                    // calculate average_potential
                    average_potential = sum_potential / counter_potential;
                    // print avergae potential
                    Console.WriteLine($"Average potential is: {average_potential}");
                }
                // increment table counter
                counter_table++;
            }
            return average_potential;
        }
        #endregion
    }
}
