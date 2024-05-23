//These first commands initialize the GUI to input image/channel information
#@ String (label="Excel File Path", description="results excel file path") PATH
#@ String (label="Channel #1 Name", description="channel 1 name") chan1
#@ String (label="Channel #2 Name", description="channel 2 name") chan2
#@ String (label="Channel #3 Name", description="channel 3 name") chan3
	
#@ Integer (label="Channel #1 Contrast/Threshold Minimum") chan1Min
#@ Integer (label="Channel #2 Contrast/Threshold Minimum") chan2Min
#@ Integer (label="Channel #3 Contrast/Threshold Minimum") chan3Min

#@ Integer (label="Channel #1 Contrast/Threshold Maximum") chan1Max
#@ Integer (label="Channel #2 Contrast/Threshold Maximum") chan2Max
#@ Integer (label="Channel #3 Contrast/Threshold Maximum") chan3Max
	
#@ Integer (label="Set scale in pixels to 100 um", value=135) bandVal

close("*") //closes all existing image windows
roiManager("reset");

input = getDirectory("Choose an input directory"); //opens desktop and prompts user to select an input directory for images to be analyzed
output = getDirectory("Choose an output directory"); //opens desktop and prompts user to select an output directory for analyzed images/results file
processFolder(input); //runs the processFolder function on the selected input directory

function processFolder(dir) 
{
	fileList = getFileList(dir); //instantiates a list of files in the input directory
	channelNames = newArray(chan1, chan2, chan3); //creates an array of channel names
	channelMin = newArray(chan1Min, chan2Min, chan3Min); //creates an array of channel threshold/contrast minimums
	channelMax = newArray(chan1Max, chan2Max, chan3Max); //creates an array of channel threshold/contrast maximums
	
	for (i=0; i<fileList.length; i++) //iterates through list of input directory files
		{
    			if(endsWith(fileList[i], ".tif")) //selects for .tif files
    	        { 
    		    	processImage(dir, output, fileList[i], channelNames, channelMin, channelMax, bandVal); //runs the measurePS function on the currently selected image
				}
				
				if(isOpen("ROI Manager")) //closes ROI manager prior to analysis of next image 
				{
	     			selectWindow("ROI Manager");
	     			run("Close");
  				}	
    	}
}

function processImage(inputFolder, output, file, chanNames, chanMins, chanMaxs, bandV) 
{
	open(inputFolder + file); //opens image file selected in processFolder
	scaledist = bandV;
	title = getTitle(); //sets title to the current image title
	coor1 = getWidth()/2; //sets coor1 to half of the selected image's width
	coor2 = getHeight()/2; //sets coor1 to half of the selected image's height
	excTitle = replace(title," ","-"); //title string (spaces replaced with hyphens)
	altThresh = ""; //variable instantiated as empty string for later use
	run("Set Scale...", "distance="+scaledist+" known=100 unit=micron");
	print("string");
	run("Split Channels"); //splits image into separate channels
	c = 0;
	
	for(l=0; l<chanNames.length; l++) //iterates through the array of image names
	{
		if(chanNames[l] == "DAPI") //adjusts image using the threshold function iff the selected window is labeled DAPI
		{
			selectWindow("C"+(l+1)+"-"+title);
			rename("DAPI");
			selectWindow("DAPI");
			threshold("DAPI", chanMins[l], chanMaxs[l], true);
		}
		else
		{
			if(c==0) //adjusts image using the threshold function iff the selected window is the first non-DAPI channel
			{
				selectWindow("C"+(l+1)+"-"+title);
				channel1 = chanNames[l]+"";
				rename(channel1);
				run("Duplicate...", "title="+channel1+"dupCol");
				threshold(channel1, chanMins[l], chanMaxs[l], false);
				tempCh1 = chanNames[l];
			}
			else //adjusts the remaining non-DAPI channel using the threshold function
			{
				selectWindow("C"+(l+1)+"-"+title);
				channel2 = chanNames[l]+"";
				rename(channel2);
				run("Duplicate...", "title="+channel2+"dupCol");
				threshold(channel2, chanMins[l], chanMaxs[l], false);
				tempCh2 = chanNames[l];
			}
			c++;
		}
	}
	
	drawROI("DAPI", coor1, coor2, 0, 100); //generates first ROI from coordinates in the middle of the window using the drawROI function
	addROIs(coor1, coor2);
	
	if("REDO"==toUpperCase(getString("Type REDO to mark the file for re-analysis or hit ENTER ", ""))) //acts as a macro ender
	{
		roiManager("reset");
		close(file);
		run("Fresh Start");
		close("*");
		run("Dispose All Windows", "/all image non-image");
	}
	
	var roiArray;
	var firstBandArray;
	var secBandArray;
	var thirdBandArray;
	var fourthBandArray;
	var fifthBandArray;
	
	selectConcentricROIs();
	
	run("Read and Write Excel", "file_mode=read_and_open file=["+PATH+"]");
	/* Following three commands measure parameters specified in "set measurements" from the specified window*/
	measureROI("DAPI"); 
	measureROI(channel1);
	measureROI(channel2);
	
	colocalize(channel1+"dupCol", channel2+"dupCol", channel1, channel2); //Performs colocalization analysis (coloc2 and area overlap) on the specified channels

	run("Read and Write Excel", "file_mode=queue_write dataset_label="+excTitle);
	run("Read and Write Excel", "file_mode=write_and_close");
	
	roiManager("reset");
	close(file);
	run("Close All");
	run("Fresh Start");
}

function threshold(win, min, max, thresh)
{
	if(thresh)
	{
		selectWindow(win);
		setThreshold(min, max, "raw");
		getThreshold(lower, upper);
		setOption("BlackBackground", true);
		
		changeThreshold = toUpperCase(getString("Current threshold: "+lower+". Auto threshold instead? (Y/N)", "N"));
		if(changeThreshold == "Y")
		{
			selectWindow(win);
			setAutoThreshold("Default dark");
			getThreshold(lower, upper);
			altThresh = " changed to "+toString(lower);
		}
		
		run("Convert to Mask");
		run("16-bit");	
		run("Smooth");
	}
	
	else
	{
		selectWindow(win);
		run("Brightness/Contrast...");
		setMinAndMax(min, max);
		run("Apply LUT");
		run("Subtract Background...", "rolling=50");
	}
}

function drawROI(wind, c1, c2, first, bV)
{
	run("ROI Manager...");
	selectWindow(wind);
	setTool("wand");
	doWand(c1, c2);
	roiManager("add");
	roiManager("select", 6*first);
	roiManager("rename", "ROI");
	Roi.setStrokeColor("red");
	
	run("Make Band...", "band="+bV);
	roiManager("add");
	roiManager("select", 6*first+1);
	roiManager("rename", "0-100");
	Roi.setStrokeColor("red");
	
	run("Make Band...", "band="+bV);
	roiManager("add");
	roiManager("select", 6*first+2);
	roiManager("rename", "100-200");
	Roi.setStrokeColor("red");
	
	run("Make Band...", "band="+bV);
	roiManager("add");
	roiManager("select", 6*first+3);
	roiManager("rename", "200-300");
	Roi.setStrokeColor("red");
	
	run("Make Band...", "band="+bV);
	roiManager("add");
	roiManager("select", 6*first+4);
	roiManager("rename", "300-400");
	Roi.setStrokeColor("red");
	
	run("Make Band...", "band="+bV);
	roiManager("add");
	Roi.setStrokeColor("red");
	roiManager("rename", "400-500");
	
	roiManager("deselect");
	//roiManager("show none");
	
	choice = toUpperCase(getString("Would you like to re-smooth(S) the image, change coordinates(C), or proceed(N)?", "N"));
	
	while(choice != "N")
	{
		if(choice == "S")
			{
			smoothAgain(wind);
			roiManager("select", 6*first+1);
			}
		else if(choice == "C")
			{
			changeCoord(wind, first, bV);
			roiManager("select", 6*first+1);
			}
		choice = toUpperCase(getString("Would you like to re-smooth(S) the image, change coordinates(C), or proceed(N)?", "N"));
	}
	
	roiManager("show none");
	
}

function addROIs(c1, c2)
{
	addROI = toUpperCase(getString("Would you like to add another ROI(Y) or proceed(N)?","N")); 
	nROI = 1; //acts as counter for while loop to generate different ROI starting coordinates
	
	while(addROI == "Y") //allows user to add as many ROIs as required
	{
		c3 = round(1600-c1/nROI);
		c4 = round(1200-c2/nROI);
		drawROI("DAPI", c3, c4, nROI, 100);
		nROI++;
		addROI = toUpperCase(getString("Would you like to add another ROI(Y) or proceed(N)?","N"));
	}
}

function selectConcentricROIs()
{
	numROIs = roiManager("count")*(1/6);
	roiManager("show all");
	print(numROIs);
	roiArray = newArray(numROIs);
	for(i=0; i<roiArray.length; i+=1)
	{
		roiArray[i] = 6*i;
	}
	
	firstBandArray = newArray(numROIs);
	for(i=0; i<firstBandArray.length; i+=1)
	{
		firstBandArray[i] = 6*i+1;
		print(firstBandArray[i]);
	}
	
	secBandArray = newArray(numROIs);
	for(i=0; i<secBandArray.length; i+=1)
	{
		secBandArray[i] = 6*i+2;
	}
	
	thirdBandArray = newArray(numROIs);
	for(i=0; i<thirdBandArray.length; i+=1)
	{
		thirdBandArray[i] = 6*i+3;
	}
	
	fourthBandArray = newArray(numROIs);
	for(i=0; i<fourthBandArray.length; i+=1)
	{
		fourthBandArray[i] = 6*i+4;
	}
	
	fifthBandArray = newArray(numROIs);
	for(i=0; i<fifthBandArray.length; i+=1)
	{
		fifthBandArray[i] = 6*i+5;
	}
	
	if(numROIs > 1)
	{
		roiManager("select", roiArray);
		roiManager("combine");
		roiManager("add");
		roiManager("deselect");
	
		roiManager("select", firstBandArray);
		roiManager("combine");
		roiManager("add");
		roiManager("deselect");
		
		roiManager("select", secBandArray);
		roiManager("combine");
		roiManager("add");
		roiManager("deselect");
	
		roiManager("select", thirdBandArray);
		roiManager("combine");
		roiManager("add");
		roiManager("deselect");
	
		roiManager("select", fourthBandArray);
		roiManager("combine");
		roiManager("add");
		roiManager("deselect");
	
		roiManager("select", fifthBandArray);
		roiManager("combine");
		roiManager("add");
		roiManager("deselect");
	}
	
	roiManager("select", firstBandArray);
	roiManager("OR");
	roiManager("add");
	roiCount = roiManager("count")-1;
	roiManager("select", roiCount);
	roiManager("rename", " 0-100");
	roiManager("deselect");
	
	xOrArray = Array.concat(roiArray,secBandArray);
	roiManager("select", xOrArray);
	roiManager("XOR");
	setBatchMode("show");
	roiManager("add");
	roiCount = roiManager("count")-1;
	roiManager("select", roiCount);
	roiManager("rename", " 100-200");
	roiManager("deselect");
	
	xOrArray = Array.concat(firstBandArray,thirdBandArray);
	roiManager("select", xOrArray);
	roiManager("XOR");
	setBatchMode("show");
	roiManager("add");
	roiCount = roiManager("count")-1;
	roiManager("select", roiCount);
	roiManager("rename", " 200-300");
	roiManager("deselect");
	
	xOrArray = Array.concat(secBandArray,fourthBandArray);
	roiManager("select", xOrArray);
	roiManager("XOR");
	setBatchMode("show");
	roiManager("add");
	roiCount = roiManager("count")-1;
	roiManager("select", roiCount);
	roiManager("rename", " 300-400");
	roiManager("deselect");
	
	xOrArray = Array.concat(thirdBandArray,fifthBandArray);
	roiManager("select", xOrArray);
	roiManager("XOR");
	setBatchMode("show");
	roiManager("add");
	roiCount = roiManager("count")-1;
	roiManager("select", roiCount);
	roiManager("rename", " 400-500");
	roiManager("deselect");
	
	run("Select None");
}

function smoothAgain(windo)
{
		run("Select None");
		selectWindow(windo);
		run("Smooth");
}
	
function changeCoord(w, num, bValue)
{
	if("A"==toUpperCase(getString("Would you like to generate the new ROI automatically(A) or manually(M)?", "A")))
	{
		setTool("wand");
		waitForUser("Please select the center point for the new ROI. Click OK when done");
		okROI = getBoolean("Click YES if ROI is acceptable or NO to return to previous window.");
	}
	else
	{
		setTool("freehand");
		waitForUser("Please outline the new ROI. Click OK when done");
		okROI = getBoolean("Click YES if ROI is acceptable or NO to return to previous window.");
	}
	if(okROI)
	{
		roiCount = roiManager("count");
		roiManager("add");
		roiManager("select", roiCount);
		roiManager("rename", "ROI");
		run("Make Band...", "band="+bValue);
		roiManager("add");
		roiManager("select", roiCount+1);
		roiManager("rename", "0-100");
		run("Make Band...", "band="+bValue);
		roiManager("add");
		roiManager("select", roiCount+2);
		roiManager("rename", "100-200");
		run("Make Band...", "band="+bValue);
		roiManager("add");
		roiManager("select", roiCount+3);
		roiManager("rename", "200-300");
		run("Make Band...", "band="+bValue);
		roiManager("add");
		roiManager("select", roiCount+4);
		roiManager("rename", "300-400");
		run("Make Band...", "band="+bValue);
		roiManager("add");
		roiManager("select", roiCount+5);
		roiManager("rename", "400-500");
		
		roiCount = roiManager("count");
		roiManager("select", roiCount-7);
		roiManager("delete");
		roiManager("select", roiCount-8);
		roiManager("delete");
		roiManager("select", roiCount-9);
		roiManager("delete");
		roiManager("select", roiCount-10);
		roiManager("delete");
		roiManager("select", roiCount-11);
		roiManager("delete");
		roiManager("select", roiCount-12);
		roiManager("delete");
		roiManager("deselect");
		
		//roiManager("show all");
	}
}

function measureROI(window)
{
	count = roiManager("count");
	
	selectWindow(window);
	roiManager("show none");
	wait(250);
	roiManager("select", count-5);
	Roi.setStrokeColor("green");
	wait(250);
	roiManager("Measure");
	Roi.setStrokeColor("red");
	roiManager("deselect");
	
	selectWindow(window);
	wait(250);
	roiManager("select", count-4);
	Roi.setStrokeColor("green");
	wait(250);
	roiManager("Measure");
	Roi.setStrokeColor("red");
	roiManager("deselect");
	
	selectWindow(window);
	wait(250);
	roiManager("select", count-3);
	Roi.setStrokeColor("green");
	wait(250);
	roiManager("Measure");
	Roi.setStrokeColor("red");
	roiManager("deselect");
	
	selectWindow(window);
	wait(250);
	roiManager("select", count-2);
	Roi.setStrokeColor("green");
	wait(250);
	roiManager("Measure");
	Roi.setStrokeColor("red");
	roiManager("deselect");
	
	selectWindow(window);
	wait(250);
	roiManager("select", count-1);
	Roi.setStrokeColor("green");
	wait(250);
	roiManager("Measure");
	Roi.setStrokeColor("red");
	roiManager("deselect");
	
	selectWindow(window);
	run("Convert to Mask");
	run("8-bit");
	copyPartToRes(window, count-5, 1);
	copyPartToRes(window, count-4, 2);
	copyPartToRes(window, count-3, 3);
	copyPartToRes(window, count-2, 4);
	copyPartToRes(window, count-1, 5);
	wait(1000);
	
	selectWindow(window);
	combArray = Array.concat(roiArray,firstBandArray,secBandArray,thirdBandArray,fourthBandArray,fifthBandArray);
	roiManager("select", combArray);
	run("Flatten");
	
	if(window == "DAPI")
	{
		saveAs("Tiff",output+altThresh+window+"-"+title);
		run("Close");
	}
	else
	{
		saveAs("Tiff",output+window+"-"+title);
		setLocation(100, 100, 40, 30);
	}
	roiManager("deselect");
}

function colocalize(win1, win2, mod1, mod2)
{
	count = roiManager("count");
	conCol = getBoolean("Would you like to measure colocalization within the ROI(s)?");
	numROI = roiManager("count");
	
	if(conCol)
	{
		wait(1000);
		selectWindow(mod1);
		wait(1000);
		run("16-bit");
		setAutoThreshold("Default dark");
		
		wait(1000);
		selectWindow(mod2);
		wait(1000);
		run("16-bit");
		setAutoThreshold("Default dark");
		
		copyPartToRes(mod1, count-3, 1);
		copyPartToRes(mod2, count-3, 1);
		copyPartToRes(mod1, count-1, 2);
		copyPartToRes(mod2, count-1, 2);
		
		selectWindow(mod1);
		run("Convert to Mask");
		run("Subtract...", "value=254");
		
		selectWindow(mod2);
		run("Convert to Mask");
		run("Subtract...", "value=254");
		
		imageCalculator("Add create", mod1, mod2);
		selectImage("Result of "+ mod1);
		
		setThreshold(2, 255);
		run("Convert to Mask");
		copyPartToRes("Result of "+ mod1, count-3, 1);
		copyPartToRes("Result of "+ mod1, count-1, 2);
		
		roiManager("deselect");
		selectWindow("DAPI");
		roiManager("select", count-3);
		run("Create Mask");
		colocAlgorithm(win1, win2, 100);
		roiManager("deselect");
		selectWindow("Mask");
		close();
		
		selectWindow(win1);
		roiManager("select", count-1);
		run("Create Mask");
		colocAlgorithm(win1, win2, 200);
		roiManager("deselect");
		selectWindow("Mask");
		close();
		
		run("Dispose All Windows", "/all image non-image");
	}
	
	else
	{
		run("Close All");
		run("Dispose All Windows", "/all image non-image");
		roiManager("reset");
	}
}

function colocAlgorithm(win1, win2, band)
{
	run("Coloc 2", "channel_1="+win1+" channel_2="+win2+" roi_or_mask=Mask threshold_regression=Bisection spearman's_rank_correlation costes'_significance_test psf=3 costes_randomisations=10");
	
	s = split(getInfo("log"),'\n');
	print(s[0]);
	
	pcc = split(s[s.length-10]);
	Array.print(pcc);
	pccVal = pcc[pcc.length-1]+"";
	print(pccVal);
	
	src = split(s[s.length-7]);
	Array.print(src);
	srcVal = src[src.length-1]+"";
	print(srcVal);
	
	cost = split(s[s.length-4]);
	Array.print(cost);
	costVal = cost[cost.length-1]+"";
	print(costVal);
	
	Table.create("LogTable");
	selectWindow("LogTable");
	Table.set("PCC", 0, pccVal); 
	Table.get("PCC", 0);
	Table.set("SRC", 0, srcVal); 
	Table.get("SRC", 0);
	Table.set("Costes", 0, costVal);
	Table.get("Costes", 0);
	Table.update;
	
	IJ.renameResults("Results", "temp");
	IJ.renameResults("LogTable","Results");
	pccStat = getResult("PCC", 0);
	srcStat = getResult("SRC", 0);
	costStat = getResult("Costes", 0);
	IJ.renameResults("temp", "Results");
	setResult("Pearson's CC", nResults, pccStat);
	setResult("Spearman's CC", nResults-1, srcStat);
	setResult("Costes", nResults-1, costStat);
	setResult("Label", nResults-1, "Stats: "+channel1+" vs "+channel2+"-"+band);
}

function copyPartToRes(mod, roiNum, band)
{
	selectWindow(mod);
	
	roiManager("select", roiNum);
	Roi.setStrokeColor("green");
	wait(250);
	
	run("Analyze Particles...", "size=0-Infinity circularity=0.00-1.00 summarize");
	selectWindow("Summary");
	IJ.renameResults("Results", "temp");
	IJ.renameResults("Summary","Results");
	channelParticleArea = getResult("Total Area", 0);
	channelParticleIntDen = getResult("IntDen", 0);
	channelParticleMean = getResult("Mean", 0);
	channelParticlePerim = getResult("Perim.", 0);
	channelParticlePercArea = getResult("%Area", 0);
	IJ.renameResults("Results","Auxillary");
	IJ.renameResults("temp", "Results");
	setResult("Area", nResults, channelParticleArea);
	updateResults();
	wait(1000);
	setResult("IntDen", nResults-1, channelParticleIntDen);
	
	if(mod.contains("Result of"))
	{
		setResult("Label", nResults-1, "Particle Analysis: "+channel1+" vs "+channel2+"-"+(band*100-100)+"-"+band*100);
	}
	else
	{
		setResult("Label", nResults-1, "Particle Analysis: "+mod+"_"+(band*100-100)+"-"+band*100);
	}
	
	setResult("Mean", nResults-1, channelParticleMean);
	setResult("Perim.", nResults-1, channelParticlePerim);
	setResult("%Area", nResults-1, channelParticlePercArea);
	
	close("Auxillary");
	updateResults();
	
	roiManager("deselect");
	
}
