__author__ = 'fturner and tjeans'
import xlrd
import os
import math
import OrcFxAPI
import sys
import win32com.client as win32
import time
import multiprocessing
#import Tkinter
#import threading
import numpy as np
from random import randint
##from scipy.optimize import curve_fit as fit
##from scipy.stats import gumbel_r, norm

# np.set_printoptions(threshold=10)
# np.set_printoptions(threshold=np.nan)



# This function just pulls the time history list from Orcaflex.
def Wave_Screen(xls_location):
    
    # initalize some variables
    start_time = time.time()
    
    c_count = 2 #changed from 1 on April 4, 2016
    attribs_per_model = False
    
    overlap, wavecritperc, ts, inp_folder, out_folder, simlen, t_wind, atts, \
             locs, n_seeds, file_list, max_gap, orcina_wb, pp_worksh,\
             resultsName, pp_wb, req, ofo, wrt_local, rand_seed,\
             scrn_hghts = Get_Inputs(xls_location)
             
             
    n_seeds = int(n_seeds)
    #req_results = [mean1, min1, max1, min_mode, max_mode, median, min_median95, max_median95, raw_data]
    avg_req = req[0]
    min_req = req[1]
    max_req = req[2]
    min_mode_req = req[3]
    max_mode_req = req[4]
    median_req = req[5]
    min_med_req = req[6]
    max_med_req = req[7]
    raw_data_req = req[8]
    
    ppMinMaxList = Get_Post_Processing(pp_wb, pp_worksh)
    wb = win32.GetObject(orcina_wb)
    # Open a sheet
    sheet1 = wb.Worksheets(resultsName)
    #the above line of code will be sending the results into the XLSM calling this Python file
    
    print ("Starting irregular analysis!")
    
    # If there are no attributes than it is intended to use the files as they are without manipulation in a one for one scenario.
    if len(atts) == 0:
        atts = []
        for i in file_list:
            atts.append(False)
        ofo = 1
        attribs_per_model = True
        print ("Using the attributes that are set in the dat files in a one for one scenario.")
        
    if ofo == 1 and len(file_list) != len(atts):
        print ("You have checked one for one but do not have the same amount of attributes as files!")
        print ("Either get rid of one for one or balance the number of files with the number of attributes to be run.")
        exit
        
    #Loop through the base file list
    current_row = 0
    for file_idx, filenme in enumerate(file_list):
        input_list=[]
        
        # print the filenames to the results spreadsheet
        sheet1.Cells(12+file_idx,1).Value = filenme
        
        # concatenate file location
        fn = os.path.join(inp_folder, filenme)
        
        # Find the wave direction for each file
        model = OrcFxAPI.Model(fn)
        wd = model.environment.WaveDirection
        
        #convert the local list of coordinates to a global one for each file.
        glob_locs=[]
        for l in locs:
            if l[0] == 'Global':
                xpos = 0
                ypos = 0
            else:
                xpos = model[l[0]].InitialX
                ypos = model[l[0]].InitialY
                hdng = model[l[0]].InitialHeading
                
            # Find global locations from locs
            if wrt_local:
                globx = xpos + l[1] * math.cos(math.radians(hdng)) + l[2] * math.sin(math.radians(hdng))
                globy = ypos - l[1] * math.sin(math.radians(hdng)) + l[2] * math.cos(math.radians(hdng))
            else:
                globx = xpos + l[1]
                globy = ypos + l[2]
                
            glob_locs.append([globx,globy])
            
        # We want to limit the amount of data in any given single time history as large time histories will crash the program
        trgchunklen = 800 # (s)        
        num_chunks = math.trunc(simlen/trgchunklen)
        if num_chunks == 0:
            num_chunks = 1
        chunklen = simlen / num_chunks
        chunklen = int(math.trunc(chunklen))
        model = ClearModel(model)
        for an, att in enumerate(atts):
            # going to now specify an attribute set for every file if all combinations or 1 attribute for every file
            if ofo == 1: #represents True
                # if the file number is different from the attribute number than the combination is skipped.
                if an != file_idx:
                    continue
                elif attribs_per_model:
                    sheet1.Cells(12+file_idx,2).Value = 'Attributes as per model'
                else:
                    #fill in the Attribute set into the column B next to the filename in Column A
                    #print 'atts header: ', 'Hs='+str(a[1]) +' Tp=' + str(a[2]) +' Gamma='+ str(a[3])
                    sheet1.Cells(12+file_idx,2).Value = 'Hs='+str(att[1]) +' Tp=' + str(att[2]) +' Gamma='+ str(att[3]) # 0=name, 1=Hs, 2=Tp, 3=gamma 
                    
            # RANDOM WAVE SEED ASSIGNMENT
            if rand_seed:
                wave_seed_int = randint(100, 9999999)
                
                print (("You have chosen to use a random wave seed: " + str(wave_seed_int)))
                if attribs_per_model:
                    model.environment.WaveSeed = wave_seed_int
                else:
                    model, wave_seed_int = SetupModel(model, att, wd, wave_seed_int)
                    
            else:
                
                if attribs_per_model:
                    print ("You have chosen to use the wave seed stored in the existing data files")
                    wave_seed_int = model.environment.WaveSeed
                else:
                    print ("You have chosen to use a repeatable wave seed of 123456")
                    wave_seed_int = 123456
                    model, wave_seed_int = SetupModel(model, att, wd, wave_seed_int)
                    
                    
            # 0=name, 1=Hs, 2=Tp, 3=gamma
            g = model.general
            g.DynamicsSolutionMethod = 'Implicit time domain'
            g.ImplicitUseVariableTimeStep = 'No'
            g.ImplicitConstantTimeStep = ts
            g.TargetLogSampleInterval = ts
            g.NumberOfStages = 1
            g.StageDuration[0] = 10.0
            # need to add a little to make sure we don't go outside the time window with the time history command.
            g.StageDuration[1] = simlen * n_seeds + 1  # So this file is actually equivalent to x number of seeds long.
            print (("Screening wave data...wave direction = " + str(wd) + ", attribute = " + str(att)))
            model.RunSimulation()
            filename3 = out_folder + "\\" + 'BlankTestModel-file=' + str(filenme) + '-att' + str(att)+'.sim'
            model.SaveSimulation(filename3)
            
            for ln, l in enumerate(glob_locs):
                
                # print the column headers and filenames to the result spreadsheets.
                # Make sure the headers only print once for the first file only.
                if file_idx == 0:
                    # For every attribute/global location we print a column header for each item in the requested results
                    # ppMinMaxList is a list of rows from the post processing worksheet.
                    for r, req_res in enumerate(ppMinMaxList):
                        if r == 0:
                            sheet1.Cells(9,c_count+1).Value = 'Global X: ' + str(l[0]) + ' , Global Y: ' + str(l[1])#'global location' #dynamic
                            if ofo == 0: #False
                                sheet1.Cells(8,c_count+1).Value = 'Hs='+str(att[1]) +' Tp=' + str(att[2]) +' Gamma='+ str(att[3]) # 0=name, 1=Hs, 2=Tp, 3=gamma
                                #sheet1.Cells(9,c_count+1).Value = 'Global X: ' + str(l[0]) + ' , Global Y: ' + str(l[1])#'global location' #dynamic
                                
                                
                        #start a loop to go through each variable since there could be more than one
                        if ',' in req_res[8]: #confirms if there is a list of variables
                            var_items = req_res[8].split(',')
                            var_items = [x.strip() for x in var_items]
                            associated = 'Extreme Value'
                            
                            for variable_item in var_items: #this would be a list if commas are found
                                c_count = Write_HDNGs(req, associated, c_count, n_seeds, req_res, variable_item, sheet1)
                                associated = 'Associated Value'
                                
                        else: #not a 'Max Associated' or 'Min Associated' command
                            associated = 'Extreme Value'
                            c_count = Write_HDNGs(req, associated, c_count, n_seeds, req_res, req_res[8], sheet1)
                            
                            
                # loop through the seeds
                for n in range (n_seeds):
                    seedList = [] #reset it
                    start = simlen * n
                    end = simlen * (n + 1)
                    
                    max_wave_height = 0 # This is where the max wave height will be initiated each time.
                    all_crit_times = []
                    
                    start = int(math.trunc(start))
                    end = int(math.trunc(end))
                    
                    #print scrn_hghts
                    #print len(scrn_hghts)
                    
                    if len(scrn_hghts) == 0:
                        for t in range(start, end, chunklen):
                            #elev = Get_Elev_TH(start, end, chunklen, t, overlap, n, simlen, lasttime, l, model.environment)
                            elev = Get_Elev_TH(start, end, chunklen, t, overlap, l, model.environment)
                            tp_only = Get_PeaksTroughsList(elev) #pass the list of index & elevations from OrcaFlex into this function
                            mwh = Get_MaxWaveHeight(tp_only) # For each peaks and trough list find the max wave height and update (function)
                            if max_wave_height <= mwh:
                                max_wave_height = mwh
                                
                    # if searching for a design wave than set it here.
                    else:
                        max_wave_height = scrn_hghts[an]
                        
                    for t in range(start, end, chunklen):
                        #elev = Get_Elev_TH(start, end, chunklen, t, overlap, n, simlen, lasttime, l, model.environment)
                        elev = Get_Elev_TH(start, end, chunklen, t, overlap, l, model.environment)                        
                        tp_only = Get_PeaksTroughsList(elev) #pass the list of index & elevations from Orcaflex into this function
                        
                        # crit_times is a list of lists [[index1, height1], [index2, height2],....] 
                        crit_times = Get_ScreenResults(wavecritperc, max_wave_height, tp_only, ts, t_wind)
                        # I think this should be when t=0 only.
                        if t == start:
                            #start index for current seed
                            start_index = t / ts
                            
                        else:
                            start_index = (t - overlap) / ts
                            
                        for c in range(len(crit_times)):
                            crit_times[c][0] = crit_times[c][0] + start_index
                            
                        all_crit_times.extend(crit_times)
                        
                    timedur = Times_Interest(all_crit_times, max_gap, ts) #[1,3], [7,5].... 1= start time, 3 = duration
                    
                    if scrn_hghts:
                        timedur = timedur[:1]
                        
                        
                    # Assemble a list of the following inputs:
                    # [[fn, a, wd, timedur, ts, out_folder, ppMinMaxList, identity list >>>[filenumber, attribute number, location number]
                    # each seed is multiprocessed with all time envelopes together so that the overall min or max value for each seed can be calculated.
                    input_list.append([fn, att, wd, timedur, ts, out_folder, ppMinMaxList, [file_idx, an, ln], wave_seed_int, attribs_per_model])
                    #print input_list
                    
                    
                    
        print (("Wave data is pre-screened for: " + str(fn)))
        # Might have a poolsize that is smaller than the number of processors minus 1.
        # Leave 1 processor for doing other work.
        free_processors = 1
        pool_size = multiprocessing.cpu_count() - free_processors
        
        # BEGINNING OF MULTIPROCESSING
        
        # The list then needs to be subdivided into chunks of info for each individual process to work on.
        # Create a list that we can pass through with pool.map
        multi_chunks=[]
        
        if len(input_list) > multiprocessing.cpu_count() - free_processors:
            if pool_size < 1:
                pool_size = 1
        elif len(input_list) <= multiprocessing.cpu_count() - free_processors:
            pool_size = len(input_list)
        
        # The list will have as many elements as there are processors minus 1
        chunk_size=math.trunc(len(input_list) / pool_size)
        remainder=len(input_list) % pool_size
        rem_count = 0
        for x in range(pool_size):
            if x+remainder+1 > pool_size:
                multi_chunks.append(input_list[(x * chunk_size) + rem_count : (x * chunk_size) + rem_count + chunk_size + 1 ])
                rem_count = rem_count + 1
            else:
                multi_chunks.append(input_list[(x * chunk_size):(x*chunk_size + chunk_size)])
                
        # This is the actual multi processing code
        
        pool = multiprocessing.Pool(processes=pool_size)
        
        print ("Start multiprocessing...")
        p = pool.map_async(Get_Results, multi_chunks)
        try:
            output = p.get(0xFFFF)
            
        except KeyboardInterrupt:
            print ("Caught keyboard interrupt, terminating workers...")
            pool.terminate()
            pool.join()
            
        pool.close()
        pool.join()
        
        # Test Code for error finding:
    ##    output = []
    ##    for chunk in multi_chunks:
    ##        output.append(Get_Results(chunk))
        # Output should be a list of dictionaries.
        # Some dictionary items may be split up among other dictionaries.  This will happen if not all
        # of the seeds are passed to the same processor per file-attrib-location key.
        # We need to merge all of the dictionaries into 1 new dictionary.
        
        print ("Finished multiprocessing, now processing results...")
        
        merged_dict={}
        for seed_dict in output:
            for key, value in seed_dict.items():
                # start by checking if the key is already in the merged dict.
                if key in merged_dict:
                    #print 'key: ' + key + '  value: ' + str(value)
                    #print value[0]
                    # if the key already exists than we have to get rid of the first value in the list before appending it.
                    del value[0]
                    merged_dict[key].extend(value)
                    
                else: # key does not exist
                    merged_dict[key] = []
                    merged_dict[key].extend(value)
                    
        # print 'merged_dict: ', merged_dict
        # empty the output list for memory
        output[:] = []
        
        # convert the dictionary to an organized list
        sorted_results = []
        
        for key, value in merged_dict.items():
            #print 'key: ', key
            #print 'value: ', value
            #convert the list from a list of seeds to results: s1        s2        r1    r2       r3
            # zip(*value[1:]) this takes a list like this  [[1,2,3], [4,5,6]] > [[1,4], [2,5], [3,6]]
            temp_list = [item for item in zip(*value[1:])]        
            median_values = [str(np.median(item)) for item in temp_list]
            avg_values = [sum(item) / len(item) for item in temp_list]
            max_values = [max(item) for item in temp_list]
            min_values = [min(item) for item in temp_list]
            
            # this block of code finds the min_mode_value and the max_mode_value
            
            min_modes = []
            max_modes = []
            # least squares
            max_ls_gumb = []
            # most likely
            max_ml_gumb = []
            
            # each item is for a different result type
            for item in temp_list:
                # hist is a list of the number of occurances
                # bin_edges will be a list with one more value than hist
                # that gives the bin edges for each bar on the graph.
                #hist, bin_edges = np.histogram(list(item), bins = n_seeds)
                hist, bin_edges = np.histogram(list(item), bins = 10)
                max_indices = [i for i, x in enumerate(hist) if x == max(hist)]
                
                min_mode = bin_edges[min(max_indices)]
                max_mode = bin_edges[max(max_indices)+1]
                min_modes.append(min_mode)
                max_modes.append(max_mode)
                
                
            # make a copy of the list before we sort it
            unsorted_list = temp_list[:]
            # put a value in for every seed for each result
            temp_list = [sorted(x, reverse=False) for x in temp_list]
            
            
            all_done = False
            halfway = math.trunc(n_seeds/2)
            
            if halfway < 2:
                # just threw this in so there will always be a lower and upper median
                lower_median_values = [0.0 for x in temp_list]
                upper_median_values = [0.0 for x in temp_list]
                
            for lower_int in range(1, halfway + 1):
                upper_int = n_seeds - lower_int + 1 
                total_sum = 0
                # start the summation between lower int and (upper int - 1)
                for k in range(lower_int, upper_int):
                    # crazy factorial equation used to define hoe many combinations in total there are.
                    combo_fn = math.factorial(n_seeds) / (math.factorial(k) * math.factorial(n_seeds - k))
                    # now plug combo_fn into weird 0.5^k*0.5^(5-k) equation thingy
                    # https://onlinecourses.science.psu.edu/stat414/node/316
                    each_part = combo_fn * 0.5**k * 0.5**(n_seeds-k)
                    total_sum = total_sum + each_part
                    
                if total_sum < 0.95:
                    # because we want a confidence interval of at least 0.95 we should go back one interval
                    if lower_int > 1:
                        lower_int = lower_int - 1
                        upper_int = upper_int + 1
                    lower_median_values = [x[lower_int-1] for x in temp_list]
                    upper_median_values = [x[upper_int-1] for x in temp_list]
                    all_done = True
                    break
                    
                if all_done:
                    break
                    
            # clearing value down to just the key information
            value = value[:1]
            # append all of the lists from above into value
            
            value.extend([avg_values, min_values, max_values, min_modes, max_modes, median_values,
                          lower_median_values, upper_median_values])
                          
            value.append(unsorted_list)
            
            sorted_results.append(value)
            
        # sorting the results based on the key values.  ie. fn, loc, attribs
        sorted_results = sorted(sorted_results, key = lambda x: (x[0][0], x[0][1], x[0][2]))
        
        print ("Results processed...")
        print ("Writing results to the appropriate row in the Excel results sheet...")
        
        row_values = []
        
        for res in sorted_results:
            # checking to see if it's time to move to the next row of data or not.
            if res[0][0] > current_row:
                # print the row values to the spreadsheet
                #print ("Output row to the Excel results sheet...")
                #sheet1.Range(sheet1.Cells(12+current_row, 2), sheet1.Cells(12+current_row, len(row_values)+1)).Value = row_values
                sheet1.Range(sheet1.Cells(12+current_row, 3), sheet1.Cells(12+current_row, len(row_values)+2)).Value = row_values
                current_row = current_row + 1
                row_values = []
                
            # arbitrarily loop through the average result values just to keep count of how many there are and to index them.
            for res_item_num, res_item in enumerate(res[1]):
                # append the flat list to row values
                if avg_req == 1: #represents True
                    row_values.append(res[1][res_item_num])
                if min_req == 1:
                    row_values.append(res[2][res_item_num])
                if max_req == 1:
                    row_values.append(res[3][res_item_num])
                if min_mode_req == 1:
                    row_values.append(res[4][res_item_num])
                if max_mode_req == 1:
                    row_values.append(res[5][res_item_num])
                if median_req == 1:
                    row_values.append(res[6][res_item_num])
                if min_med_req == 1:
                    row_values.append(res[7][res_item_num])
                if max_med_req == 1:
                    row_values.append(res[8][res_item_num])
                if raw_data_req == 1:
                    row_values.extend(res[9][res_item_num])
                    
                    
        sheet1.Range(sheet1.Cells(12+current_row, 3), sheet1.Cells(12+current_row, len(row_values)+2)).Value = row_values
        wb.Save()
        print ("Moving on to next file...")
    wb.Close(True)
    print ("Analysis complete!")
    print (("Time to complete this analysis (in seconds): " + str(time.time()-start_time)))



def ClearModel(model):
    
    for i in model.objects:
##        print i.handle
##        print i.typeName
##        #print i.name
##        print i
        
        # remove lines, winches, and links first
        if i.type in (6,9,10):
            try:
                model.DestroyObject(i.name)
            except:
                pass
                
    for i in model.objects:
        # remove line types
        if i.type == 15:
            try:
                model.DestroyObject(i.name)
            except:
                pass
                
    for i in model.objects:
        # remove everything else
        if i.type not in (1,3):
            try:
                model.DestroyObject(i.name)
            except:
                pass
                
    return model



def Write_HDNGs(req, assoc, c_count, n_seeds, req_result, var_item, sht):
    
    first_req = True
    for xx, request in enumerate(req):
        if request == 0: #represents False
            continue 
        if xx == 0:
            reqstr = 'MEAN'
        elif xx == 1:
            reqstr = 'MIN'
        elif xx == 2:
            reqstr = 'MAX'
        elif xx == 3:
            reqstr = 'MIN-MODE'
        elif xx == 4:
            reqstr = 'MAX-MODE'
        elif xx == 5:
            reqstr = 'MEDIAN'
        elif xx == 6:
            reqstr = '95%-MIN-MEDIAN'
        elif xx == 7:
            reqstr = '95%-MAX-MEDIAN'
##        elif xx == 8:
##            reqstr = 'GUMBEL MOST LIKELY'
##        elif xx == 9:
##            reqstr = 'GUMBEL LEAST SQUARES'
        elif xx == 8:
            reqstr = 'RAW DATA SEED'
            
            
        if first_req:
            # advance the column count by 1
            c_count = c_count + 1
            colHeader = str(req_result[4]) + "-" + assoc + "-" + str(var_item) \
                        + "-" + str(req_result[6]) + "-" + str(req_result[5]) + "-" + reqstr
            sht.Cells(11, c_count).Value = colHeader
            
        elif xx == 8:
            for sd_num in range(n_seeds):
                c_count = c_count + 1
                sht.Cells(11, c_count).Value = reqstr + ': ' + str(sd_num + 1)
                
        else:
            c_count = c_count + 1
            sht.Cells(11, c_count).Value = reqstr
            
        first_req = False
        
    return c_count



def Get_Results(args):
    # Get Results is run one time for each unique seed case.
    # Create a dictionary to store a key representing the
    # data position in the result table and the results itself
    
    temp_result_dict = {}
    
    # [[fn, a, wd, timedur, ts, out_folder, ppMinMaxList
    
    for arg in args:
        f_name = arg[0]
        attrib = arg[1]
        w_dir = arg[2]
        t_dur = arg[3]
        t_step = arg[4]
        model_folder = arg[5]
        Requested_Results = arg[6]
        identity_list = arg[7] #[filenumber, attribute number, location number]
        w_seed = arg[8]
        att_per_mod = arg[9]
        
        #print('wave seed3: ' + str(w_seed))
        
        # convert list into a string to be used as a unique key for the seed group
        temp_res_key = ''.join(str(e) for e in identity_list)
        # if the dictionary key does not exist yet then add the key and the result identity values
        if temp_res_key not in temp_result_dict:
            temp_result_dict[temp_res_key] = []
            temp_result_dict[temp_res_key].append(identity_list)
            
        while True:
            try:
                model2 = OrcFxAPI.Model(f_name)
                if att_per_mod:
                    model2.environment.WaveSeed = w_seed
                else:
                    model2, w_seed = SetupModel(model2, attrib, w_dir, w_seed)
                
                result_set = []
                # fails represents all of the failed attempts so that we know if we are appending to a list of modifying a list.
                fails = 0
                for p, q in enumerate (t_dur):
                    
                    #we want the sim to run 40 seconds after the period of interest
                    duration = q[1] * t_step + 80
                    start_t = q[0] * t_step - 40
                    
                    #For naming convention use the [filename, atts, glob loc], start time and duration
                    filename2 = os.path.join(model_folder, (temp_res_key + '-start=' +\
                                                            str(math.trunc(start_t)) +\
                                                            '-dur=' + str(math.trunc(duration)) + '.sim'))
                    opened = False
                    
                    try:
                        model2.LoadSimulation(filename2)
                        print (("Open file to get results: " + str(filename2)))
                        opened = True
                    except:
                        print (("File does not already exist; running file: " + str(filename2)))
                        
                    if not opened:
                        
                        #we want the sim to run 20 seconds before the period of interest
                        model2.general.NumberOfStages = 1
                        model2.general.StageDuration[0] = 10.0 #stage 0 is the buildup
                        
                        model2.general.StageDuration[1] =  duration
                        model2.environment.SimulationTimeOrigin = start_t
                        
                        endit = False
                        
                        while True:
                            try:
                                model2.RunSimulation()
                                print (("Running simulation... " + str(filename2)))
                                if not model2.simulationComplete:
                                    raise Exception('Simulation did not complete...')
                                break
                            except:
                                print (("Failure running model: " + str(filename2) + "; Try cutting time step in half"))
                                if model2.general.DynamicsSolutionMethod == 'Explicit time domain':
                                    ts = model2.general.InnerTimeStep
                                    model2.general.InnerTimeStep = ts / 2
                                else:
                                    if model2.general.ImplicitUseVariableTimeStep == 'Yes':
                                        ts = model2.general.ImplicitVariableMaxTimeStep
                                        model2.general.ImplicitVariableMaxTimeStep = ts / 2
                                    else:
                                        ts = model2.general.ImplicitConstantTimeStep
                                        model2.general.ImplicitConstantTimeStep = ts / 2
                                print ((str(filename2) + "; Time step is now: " + str(ts)))
                                if ts < 0.0001:
                                    fails = fails + 1
                                    print (("Ending this simulation as it is unstable: " + str(filename2)))
                                    print ("WARNING **************************************************************************************************************************************************************")
                                    print ("Data may not be accurate as a simulation will now be missing from the set.")
                                    endit = True
                                    break
                                else:
                                    continue
                            # this gets us out of the while loop.
                            if endit:
                                break
                        if endit:
                            # going on to the next simulation as this one is unstable.
                            continue
                            
                            
                        #read in the entire set of data from Orcina Excel file
                        #Save the file as a SIM file here
                        model2.SaveSimulation(filename2)
                        print (("Saved: " + str(filename2)))
                        
                    ers = 0 # extra row index that is used to index our result_set[ers+m] list variable in addition with m.
                    
                    for m, extreme in enumerate (Requested_Results):
                        
                        additData = Requested_Results[m][6]
                        variable1 = Requested_Results[m][8] #represents the Column I which is the Variable column in Excel
                        objtype = Requested_Results[m][5]
                        command = Requested_Results[m][4]
                        line = model2[objtype]
                        
                        if command in ("Min", "Max"):
                            if additData[:9] == "ArcLength":#should be the same as the Additional Data column in Excel.
                            #Now to get the number associated with the Arclength which feeds into the below command/function
                                oeArcNum = float(additData[9:])
                                TH = line.TimeHistory(variable1, 1, OrcFxAPI.oeArcLength(oeArcNum))
                            elif additData[:4] == "Node":
                                #Now to get the number associated with the Node which feeds into the below command/function
                                TH = line.TimeHistory(variable1, 1, OrcFxAPI.oeNodeNum(Requested_Results[6][7:]))
                            elif additData[:5] == "End A":
                                TH = line.TimeHistory(variable1, 1, OrcFxAPI.oeEndA)
                            elif additData[:5] == "End B":
                                TH = line.TimeHistory(variable1, 1, OrcFxAPI.oeEndB)
                            elif additData[:9] == "Touchdown":
                                TH = line.TimeHistory(variable1, 1, OrcFxAPI.oeTouchdown)
                            else:
                                TH = line.TimeHistory(variable1, 1)
                                
                            if command == "Min":
                                #take the MIN of the list and append the min's into a list
                                if (p - fails) == 0:
                                    result_set.append(min(TH))
                                elif min(TH) < result_set[m+ers]:
                                    result_set[m+ers] = min(TH)
                                    
                            elif command == "Max":
                                #take the MAX of the list and append the max's into a list
                                if (p - fails) == 0:
                                    result_set.append(max(TH))
                                elif max(TH) > result_set[m+ers]:
                                    result_set[m+ers] = max(TH)
                                    
                        elif command in ("Range Graph Min", "Range Graph Max"):
                            if additData == '': #represents arEntireLine if left blank
                                RGR = line.RangeGraph(variable1, 1, None, OrcFxAPI.arEntireLine())
                            elif 'to' in additData:
                                to_loc = additData.index(' to ') #Get_Post_Processing function makes sure that user entered a value containing ' to ' or blank
                                # arc length start
                                al_start = float(additData[ : to_loc])
                                al_end = float(additData[to_loc+4 : ])
                                RGR = line.RangeGraph(variable1, 1, None, OrcFxAPI.arSpecifiedArclengths(al_start,al_end))
                            else:
                                print (("Error with the additional data column for command: " + str(m)))
                                
                            if command == "Range Graph Min":
                                #take the MIN of the list and append the min's into a list
                                if (p - fails) == 0:
                                    result_set.append(min(RGR.Min))
                                elif min(RGR.Min) < result_set[m+ers]:
                                    result_set[m+ers] = min(RGR.Min)
                                    
                            elif command == "Range Graph Max":
                                #take the MAX of the list and append the max's into a list
                                if (p - fails) == 0:
                                    result_set.append(max(RGR.Max))
                                elif max(RGR.Max) > result_set[m+ers]:
                                    result_set[m+ers] = max(RGR.Max)
                                    
                        elif command in ("Min Associated", "Max Associated"):
                            variable_list = [x.strip() for x in variable1.split(',')]
                            variable_results = []
                            
                            for vnum, var in enumerate(variable_list):
                                
                                if additData[:9] == "ArcLength":#should be the same as the Additional Data column in Excel.
                                #Now to get the number associated with the Arclength which feeds into the below command/function
                                    oeArcNum = float(additData[9:])
                                    TH = line.TimeHistory(var, 1, OrcFxAPI.oeArcLength(oeArcNum))
                                elif additData[:4] == "Node":
                                    #Now to get the number associated with the Node which feeds into the below command/function
                                    TH = line.TimeHistory(var, 1, OrcFxAPI.oeNodeNum(Requested_Results[6][7:]))
                                elif additData[:5] == "End A":
                                    TH = line.TimeHistory(var, 1, OrcFxAPI.oeEndA)
                                elif additData[:5] == "End B":
                                    #print 'variable: ', var
                                    TH = line.TimeHistory(var, 1, OrcFxAPI.oeEndB)
                                elif additData[:9] == "Touchdown":
                                    TH = line.TimeHistory(var, 1, OrcFxAPI.oeTouchdown)
                                else:
                                    print (("Error with additional data column for command: " + str(m)))
                                # make the Time History numpy array a list so we can get the index later (https://stackoverflow.com/questions/1966207/converting-numpy-array-into-python-list-structure)
                                TH = np.array(TH).tolist()
                                
                                
                                # want the minimum/maximum value of the first variable only
                                if vnum == 0:
                                    if command == "Min Associated":
                                        # this will take the value nearest to 0, regardless of sign convention (https://stackoverflow.com/questions/44864633/pythonic-way-to-find-maximum-absolute-value-of-list)
                                        extreme_val = min(TH, key=abs)
                                    elif command == "Max Associated":
                                        # this will take the value furthest away from 0, regardless of sign convention (https://stackoverflow.com/questions/44864633/pythonic-way-to-find-maximum-absolute-value-of-list)
                                        extreme_val = max(TH, key=abs)
                                    variable_results.append(extreme_val)
                                    # get the position which the minimum/maximum value occurs in the list; this only works because we made the numpy array a list above
                                    res_index = TH.index(extreme_val)
                                    
                                # want the associated value (associated to the first variable) of remaining variables
                                else:
                                    variable_results.append(TH[res_index])
                                    
                            # q takes into account any simulations that failed so that we make sure we only extend the list on the first simulation.            
                            if (p - fails) == 0:
                                result_set.extend(variable_results)
                            elif variable_results[0] < result_set[m+ers] and command == "Min Associated":
                               for vn, var, in enumerate(variable_list):
                                    result_set[m+ers+vn] = variable_results[vn]
                            elif variable_results[0] > result_set[m+ers] and command == "Max Associated":
                                for vn, var, in enumerate(variable_list):
                                    result_set[m+ers+vn] = variable_results[vn]
                                    
                                    
                            # normally we have 1 result per row in the post processing sheet
                            # then we can index the results with variable "m" which just counts those rows.
                            # however for associated results we can have len(variable_list) - 1 extra results for each row.
                            # need to keep track of this.
                            ers = ers + len(variable_list) - 1
                        else:
                            print ("Invalid Extraction Command!")
                            raise Exception("INVALID EXTRACTION COMMAND")
                            
                temp_result_dict[temp_res_key].append(result_set)
                #print temp_result_dict
                break
                
            except OrcFxAPI.DLLError as e:
                print (e)
                if e.status == OrcFxAPI.stLicensingError:
                    print ("OrcaFlex license error...sleeping for 5s...")
                    time.sleep(5)
                    
                elif e.status == OrcFxAPI.stStaticsFailed:
                    print (("Statics failed for: " + str(filename2)))
                    break
                    
                elif e.status == OrcFxAPI.stFileReadError or e.status == OrcFxAPI.stFileWriteError or e.status == OrcFxAPI.stFileNotFound:
                    print (("File read or write error for file: " + str(filename2)))
                    print ("Sleeping for 60s")
                    time.sleep(60)
                    break
                    
                elif e.status == OrcFxAPI.stOperationCancelled:
                    print (("Simulation failed; most likely a timeout error! Either increase the timeout limit or improve the simulation speed for: " + str(filename2)))
                    break
                    
                else:
                    print (("Getting error other than license error or statics failed here. This could be a bad configuration limit setting. Ultimately this will end in an unstable run." + str(filename2) + " failed."))
                    break
                    
    return temp_result_dict 



def SetupModel(m, att, wavedir, ws):#model and attribute
    e = m.environment
    e.NumberOfWaveTrains = 1
    e.WaveName[0] = 'Wave1' # Set name of wave train
    e.SelectedWaveTrain = e.WaveName[0] # Select the wave train
    e.WaveDirection = wavedir # Set the wave direction for the wave train
    e.WaveType = 'JONSWAP'
    
    # 0=name, 1=Hs, 2=Tp, 3=gamma
    e.WaveJONSWAPParameters  = 'Partially Specified'
    e.WaveOriginX  = 0.0
    e.WaveOriginY  = 0.0
    e.WaveTimeOrigin  = 0.0
    e.WaveHs = att[1]#attributes from XLS
    e.WaveGamma = att[3]
    e.WaveTp = att[2]
    e.WaveFrequencySpectrumDiscretisationMethod = 'Equal energy'
    
    e.UserSpecifiedRandomWaveSeeds = 'Yes'
    e.WaveSeed = ws
    
    e.WaveNumberOfComponents = 200
    
    return m, ws



def Listsum(numList):
    if len(numList) == 1:
        return numList[0]
    else:
        return numList[0] + Listsum(numList[1:])



def Get_Elev_TH(start, end, chunklen, time, overlap, l, e):
    # To account for a buffer/overlap between data sets we will adjust the seg_start and
    # seg_end values.
    
    if time + chunklen + overlap >= end: # added overlap
        seg_end = end
    else:
        seg_end = time + chunklen + overlap
        
        
    if time - overlap <= start:
        seg_start = time
    else:
        seg_start = time - overlap
        
        
    Elev_TH = e.TimeHistory('Elevation', OrcFxAPI.SpecifiedPeriod(seg_start, seg_end), OrcFxAPI.oeEnvironment(l[0],l[1],0.0))
    return Elev_TH



def Get_PeaksTroughsList(elev):
    
    troughpeaksonly = []
    
    #initialize trough and peak to the first element in the list
    trough = elev[0]
    peak = elev[0]
    
    # initialize the very first wave slope as either being up or down based on the
    # first and second element.
    if elev[1] > elev[0]:
        slope_up = True
    else:
        slope_up = False
        
    # populate the very first element in the new tp_list.
    # first value is the index, second value is the elevation
    troughpeaksonly.append([0,elev[0]])
    
    #CREATE LIST OF PEAKS AND TROUGHS ONLY
    for i, e in enumerate (elev):
        # Check the slope to determine if it's going to be a peak or trough
        if slope_up == True:
            # If the next element is higher than the current peak value than update the peak.
            if elev[i+1] >= peak:
                peak = elev[i+1]
                
            # If the elevation is less than the current peak then we know we have just passed
            # the peak.
            # We also know that we already have a trough value stored above.
            # At this point we are 1 index value ahead of the actual peak value.
            else:
                slope_up = False
                troughpeaksonly.append([i,e])
                trough = e
                
        else: # We're on the down swing
            if elev[i+1] < trough:
                trough = elev[i+1]
            else:
                slope_up = True
                troughpeaksonly.append([i,e])
                peak = elev[i+1]
        if i+3 > len(elev):
            troughpeaksonly.append([i+1,elev[i+1]]) #add the last trough/peak in the list
            break
    return troughpeaksonly



def Get_MaxWaveHeight(troughpeaksonly):
    #FIND MAX WAVE HEIGHT
    wavemaxheight = 0.0 #initialize
    for i,ie_pair in enumerate(troughpeaksonly):
        if abs(ie_pair[1] - troughpeaksonly[i+1][1]) > wavemaxheight:
                wavemaxheight = abs(ie_pair[1] - troughpeaksonly[i+1][1])
        if i+3 > len(troughpeaksonly): # we check i+3 to account for the length being one longer than the index.
            break
    return (wavemaxheight)



def Get_ScreenResults(wavecritpercent, wavemaxheight, troughpeaksonly, timestep, timewindow):
    wavecrit = wavecritpercent * wavemaxheight
    indexwindow = timewindow/timestep #This is just representing the index window
    crit_results = [] #results above the criteria in the excel sheet.
    prevwave = troughpeaksonly[0][1] #first trough or peak elevation in the list
    endofdata=False
    for n, ie_pair in enumerate(troughpeaksonly):
        if endofdata:
            break
        # first thing we need to do is to look ahead until we find our actual data window.
        for x in range(n, len(troughpeaksonly)):
            if troughpeaksonly[x][0] - troughpeaksonly[n][0] > indexwindow:
                # first value of the range is actually the n value
                # The last value of the range is actually the x-1 value when this if statement
                #is true.
                if x-n == 1:
                    current_window = troughpeaksonly [n:x+1]
                else:
                    current_window = troughpeaksonly [n:x]
                if Get_Time(current_window, wavecrit):
                    crit_results.append(Get_Time(current_window, wavecrit))
                    #print 'crit_results1: ', crit_results
                break
            elif x+1==len(troughpeaksonly) and troughpeaksonly[x][0] - troughpeaksonly[n][0] <= indexwindow:
                current_window = troughpeaksonly [n:x+1]
                endofdata=True
                if Get_Time(current_window, wavecrit):
                    crit_results.append(Get_Time(current_window, wavecrit))
                    #print 'crit_results2: ', crit_results
                break
    return (crit_results)



def Times_Interest(allcrittimes, maxgap, t_s):
    #remove dup times from the list
    #Set function does not keep the order of the list, thus a loop is used below
    newlist = []
    newlist2 = []
    for i in allcrittimes:
        if i not in newlist:
            newlist.append(i)
    newlist2 = [item[0] for item in newlist]# this will return just the indexes in the list
    
    
    #since we no longer need the elevations - we just need the times or indexes
    final_list= []    
    startpos = newlist2[0]#initialize it with the first value from the list
    lastpos = newlist2[0] 
    for i, j in enumerate (newlist2):
        
        if j-startpos > (maxgap / t_s) or newlist2[-1] == j: #put in a secondary check to see if this is the last item in the list
            
            if newlist2[-1] == j:
                if j-startpos > (maxgap / t_s): #this represents the last run through the loop
                    final_list.append([startpos, lastpos - startpos])
                    final_list.append([j, 0])
                else:
                    final_list.append([startpos, j - startpos])
                    
            else:
                final_list.append([startpos, lastpos - startpos])
                startpos = j
                lastpos = j
                
        else:
            lastpos = j
            
    #...when two or more times fall within the maxsimlen in which case we only want the first
    #time index and the duration to get the last one in that maxsimlen window
    
    #end result will be a list of [start time, duration]
    return final_list



def Get_Time(cur_win, crit):
    max_val_pair = max(cur_win, key=lambda x: x[1])
    min_val_pair = min(cur_win, key=lambda x: x[1])
    max_val = max_val_pair[1]
    max_index = max_val_pair[0]
    min_val = min_val_pair[1]
    min_index = min_val_pair[0]
    height = max_val-min_val
    
    if height >= crit:
        centroid_index = math.trunc((max_index+min_index)/2) 
        return ([centroid_index, height])
        
    return False



def Get_Inputs(file_location):
    #Get Inputs from Excel    
    workbook = xlrd.open_workbook(file_location)
    sheet = workbook.sheet_by_name(inputWS) #worksheet calling Python - may not be 'sheet1' anymore so needs to be dynamic
    maxrow = sheet.nrows
    print (("Max row: " + str(maxrow)))
    # initiate all end cell row numbers to the max row
    endcellnums = {'filename_endcellnum':[5, maxrow], 'atts_endcellnum':[8, maxrow],
                   'scrn_endcellnum':[11, maxrow], 'glblocs_endcellnum':[12, maxrow]}
                   
                   
    # find the maximum cell numbers
    for col in endcellnums:
        for row in range(3, maxrow):
            if sheet.cell(row, endcellnums[col][0]).value == '':
                endcellnums[col][1] = row
                break
                
    print (endcellnums)
    filename_endcellnum = endcellnums['filename_endcellnum'][1]
    atts_endcellnum = endcellnums['atts_endcellnum'][1]
    scrn_endcellnum = endcellnums['scrn_endcellnum'][1]
    glblocs_endcellnum = endcellnums['glblocs_endcellnum'][1]
    
    
    input_fldr = sheet.cell(3,1).value #input folder where file list will reside
    resultsWSName = sheet.cell_value(5,3)#Dynamic name of Results worksheet (where data will be copied to from Python) since it could be "Results1", Results2", etc
    pp_workbook = sheet.cell_value(5,1) #Location of XLS where Post-processing worksheet resides. The user has the option
    #to copy the info from another XLS and place into the file calling this Python file or they can leave it in an existing XLS. We just need to determine where to find the post-process info
    pp_worksheet = sheet.cell_value(6,1) #Orcina spreadsheet worksheet name since the user may change it from the default "Post-processing"
    orcina_workbook = os.path.join(sheet.cell_value(7,1), sheet.cell_value(8,1)) #Orcina spreadsheet filename (no path - needs a path)
    #print "orcina_workbook in Get_Results: ", orcina_workbook
    sim_len = sheet.cell_value(9,1) #simulation length
    seeds = sheet.cell_value(10,1) #Number of seeds
    maxgap = sheet.cell_value(11,1) #max gap
    wave_crit = sheet.cell_value(12,1)/100 #this next portion of its calculation (* wavemaxheight) will be handled in the Get_PeaksTroughsList function 
    t_s = sheet.cell_value(13,1) #time step
    time_wind = sheet.cell_value(14,1) #time window
    over_lap = sheet.cell_value(15,1)
    output_fldr = sheet.cell_value(7,1) #output folder where sim files will reside
    one_for_one = sheet.cell_value(9,3) #determines if user wants to run ALL Attributes against all DATs or if they want to run one set of Atts which is listed directly across from the DAT in Excel
    max1 = sheet.cell_value(10,3) #max checkbox
    min1 = sheet.cell_value(11,3) #min checkbox
    mean1 = sheet.cell_value(12,3) #mean checkbox
    max_mode = sheet.cell_value(13,3) #max mode checkbox
    min_mode = sheet.cell_value(14,3) #min mode checkbox
    median = sheet.cell_value(15,3) #median checkbox
    min_median95 = sheet.cell_value(16,3) #min median 95% checkbox
    max_median95 = sheet.cell_value(17,3) #max median 95% checkbox
    raw_data = sheet.cell_value(18,3) #raw data checkbox
    local = sheet.cell_value(19,3)
    random_seed = sheet.cell_value(20,3)
    
    req_results = [mean1, min1, max1, min_mode, max_mode, median, min_median95, max_median95, raw_data]
    
    filelist = sheet.col_values(5,3,filename_endcellnum) #need to know the last row
    #print filelist
    
    attslistCol1 = sheet.col_values(7,3,atts_endcellnum) #list of attributes ColH (Name)
    attslistCol2 = sheet.col_values(8,3,atts_endcellnum) #list of attributes ColI (hs)
    attslistCol3 = sheet.col_values(9,3,atts_endcellnum) #list of attributes ColJ (tp)
    attslistCol4 = sheet.col_values(10,3,atts_endcellnum) #list of attributes ColK (gamma)
    irr_wve_atts = list(zip(attslistCol1, attslistCol2, attslistCol3, attslistCol4))
    
    screen_heights = sheet.col_values(11,3,scrn_endcellnum) #list of attributes ColK (gamma)
    
    glb_locsCol1 = sheet.col_values(12,3,glblocs_endcellnum) #list of global locations
    glb_locsCol2 = sheet.col_values(13,3,glblocs_endcellnum) #list of global locations
    glb_locsCol3 = sheet.col_values(14,3,glblocs_endcellnum) #list of global locations
    glb_locs = list(zip(glb_locsCol1, glb_locsCol2, glb_locsCol3))
    
    return over_lap, wave_crit, t_s, input_fldr, output_fldr, sim_len, time_wind, irr_wve_atts, glb_locs, seeds,\
           filelist, maxgap, orcina_workbook, pp_worksheet, resultsWSName, pp_workbook, req_results, one_for_one, local,\
           random_seed, screen_heights



def Get_Post_Processing(xls_locale, ppworksh):
    #Get Inputs from Orcaflex Excel file which contains the post-processing info generated from the Instructions Wizard
    workbook = xlrd.open_workbook(xls_locale)#this could be local XLS or a separate XLS
    sheet = workbook.sheet_by_name(ppworksh)#worksheet is dynamic as it will be passed in via Excel input and user may have renamed it there
    #do a loop through the rows in this worksheet and append to a list only those that are 'Min' & 'Max' in Column E
    #We are only concerned with the first section of the worksheet (up until the second SIM file name is listed like a header)
    rownum = sheet.nrows -1#is supposed to return the number of rows from Excel
    ppList = []#pp stands for Post-Processing
    ppListFiltered = [] #reset it
    #Set a counter so that the word 'Load' can occur once but once it appears a second time, it is time for the loop to end
    #since it will signify the start of the second SIM file data which we are not interested in
    load_found = False
    for i in range(sheet.nrows):
        ppList.append(sheet.row_values(i))
    for i, j in enumerate(ppList):
        if ppList[i][4][:4] in ("Load","load"):
            if not load_found: # Should only be one set of instructions
                #print ("First instance of Load File command found in post processing instruction sheet...")
                load_found = True
            else: # Should only be one set of instructions
                print ("Second instance of Load File command found in post processing instruction sheet. There should only be one set of instructions. These additional instructions will be ignored.")
                break
                
        elif ppList[i][4] in (""):
            if not load_found:
                print ("Blank line skipped in post processing instruction sheet...")
            else:
                print ("End of post processing instructions found...")
                break
            
        elif ppList[i][4] in ("Command"):
            print ("Heading line skipped in post processing instruction sheet...")
            
        elif ppList[i][4] in ("Min","Max"):
            ppListFiltered.append(j)
            
        elif ppList[i][4] in ("Min Associated","Max Associated"):
            try:
                test_comma = ppList[i][8].index(",")
            except ValueError:
                print ("Associated command must have a minimum of two variables. Exiting Python...")
                sys.exit()
            ppListFiltered.append(j)
            
        elif ppList[i][4] in ("Range Graph Min","Range Graph Max"):
            if ppList[i][6] != "": #represents an Entire Line if left blank
                try:
                    test_to = ppList[i][6].index(" to ")
                except ValueError:
                    print ("Invalid Range Graph value from Excel. Exiting Python...")
                    sys.exit()
            ppListFiltered.append(j)
            
        elif ppList[i][4] in ("Min Associated Range Graph","Max Associated Range Graph"):
            try:
                test_comma = ppList[i][8].index(",")
            except ValueError:
                print ("Associated command must have a minimum of two variables. Exiting Python...")
                sys.exit()
            if ppList[i][6] != "": #represents an Entire Line if left blank
                try:
                    test_to = ppList[i][6].index(" to ")
                except ValueError:
                    print ("Invalid Associated Range Graph value from Excel. Exiting Python...")
                    sys.exit()
            ppListFiltered.append(j)
            
        else:
            print ("Result instruction is not recognized, talk to Fred.")
            sys.exit()
            
    return ppListFiltered



def sync(basefile,folder): #This will extract the model type and name of each OrcaFlex object from each .DAT file in Excel,
    #then put into a temp text file, then populate a list in Excel
    print ("Syncing... Please wait...")
    folder = sys.argv[3] #path
    filename = 'temp.txt'
    tempfile = os.path.join(folder, filename)
    f = open(tempfile,'w')
    basefile = sys.argv[2] #path & filename of DAT file
    ThreeDBuoys = []
    SixDBuoys = []
    Vessels = []
    
    while True:
        try:
            model = OrcFxAPI.Model(str(basefile))
            #Adding Vessels into the vessel list
            Vessels.append([i.name for i in model.objects if i.type == OrcFxAPI.otVessel])
            #Adding 3D buoys into the 3DBuoys list
            ThreeDBuoys.append([i.name for i in model.objects if i.type == OrcFxAPI.ot3DBuoy])
            #Adding 6D buoys into the 6DBuoys list
            #SixDBuoys.append([i.name for i in model.objects if i.type == OrcFxAPI.ot6DBuoy])
            break
        except OrcFxAPI.DLLError as e:
            if e.status==OrcFxAPI.stLicensingError:
                logging.info("OrcaFlex license error")
                time.sleep(5)
                continue
            else:
                logging.info("Getting error other then license error here")
                continue
                
    #Writing values to temp.txt
    #print "Starting to write to file"
    for i in Vessels[0]:
        f.write(str(i))
        f.write('\n')
        
    f.write('\n')
    
    for i in ThreeDBuoys[0]:
        f.write(str(i))
        f.write('\n')
        
    f.write('\n')
    
    """for i in SixDBuoys[0]:
        f.write(str(i))
        f.write('\n')
    
    f.write('\n')"""
    
    print ("Sync complete!")
    f.close()



def LicenceErrHandler(action):
    global licenceRetryAttempts
    if action==OrcFxAPI.lrBegin: # first call, set initial retry count.
        print ("OrcaFlex license fail; retry attempt number 1")
        time.sleep(10)  # wait a bit
        licenceRetryAttempts = 1
        return True
    elif action==OrcFxAPI.lrContinue: # subsequent retry calls land here. Choose to continue retrying or give up
        if licenceRetryAttempts>1000:
            print ("Given up trying to get an OrcaFlex license.")
            return False
        else:
            print (("OrcaFlex license fail; retry attempt number " + str(licenceRetryAttempts)))
            time.sleep(10)  # wait a bit
            licenceRetryAttempts += 1
            return True
    elif action==OrcFxAPI.lrEnd:  # this action occurs after retry attempts are stopped (successfully or not). Opportunity to reset any relevant data
        return False



if __name__ == '__main__':
    option = sys.argv[1]
    OrcFxAPI.RegisterLicenceNotFoundHandler(LicenceErrHandler)
    multiprocessing.freeze_support() # pyinstaller will not work without this!!!
    
    if option == str(1):
        sync(sys.argv[2], sys.argv[3]) #basefile and folder arguments
    else:
        xls_loc = sys.argv[1] #file that is calling this Python file which may or may not contain the post-processing data (up to the user if they want that info in current XLS or an external one)
        inputWS = sys.argv[2]#"Input"  #worksheet that is calling this Python file
        Wave_Screen(xls_loc)

