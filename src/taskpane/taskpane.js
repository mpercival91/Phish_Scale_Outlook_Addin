/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office */


Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    //document.getElementById("run").onclick = run;
	
	//listeners for changes to selectors to recalculate within interaction 
	document.getElementById ("cue1").addEventListener ('change', run);
	document.getElementById ("cue2").addEventListener ('change', run);
	document.getElementById ("cue3").addEventListener ('change', run);
	document.getElementById ("cue4").addEventListener ('change', run);
	document.getElementById ("cue5").addEventListener ('change', run);
	document.getElementById ("cue6").addEventListener ('change', run);
	document.getElementById ("cue7").addEventListener ('change', run);
	document.getElementById ("cue8").addEventListener ('change', run);
	document.getElementById ("cue9").addEventListener ('change', run);
	document.getElementById ("cue10").addEventListener ('change', run);
	document.getElementById ("cue11").addEventListener ('change', run);
	document.getElementById ("cue12").addEventListener ('change', run);
	document.getElementById ("cue13").addEventListener ('change', run);
	document.getElementById ("cue14").addEventListener ('change', run);
	document.getElementById ("cue15").addEventListener ('change', run);
	document.getElementById ("cue16").addEventListener ('change', run);
	document.getElementById ("cue17").addEventListener ('change', run);
	document.getElementById ("cue18").addEventListener ('change', run);
	document.getElementById ("cue19").addEventListener ('change', run);
	document.getElementById ("cue20").addEventListener ('change', run);
	document.getElementById ("cue21").addEventListener ('change', run);
	document.getElementById ("cue22").addEventListener ('change', run);
	document.getElementById ("premise_alignment_assessment_point_1").addEventListener ('change', run);
	document.getElementById ("premise_alignment_assessment_point_2").addEventListener ('change', run);
	document.getElementById ("premise_alignment_assessment_point_3").addEventListener ('change', run);
	document.getElementById ("premise_alignment_assessment_point_4").addEventListener ('change', run);
	document.getElementById ("premise_alignment_assessment_point_5").addEventListener ('change', run);
  }
});

export async function run() {
	 // Get a reference to the current message
	var item = Office.context.mailbox.item;
	
	var final_premise_alignment_score = 0;
	//Get 1st result and convert to a score
	var assessment_point_1_selection = document.getElementById("premise_alignment_assessment_point_1");
	var assessment_point_1_selection_text = assessment_point_1_selection.options[assessment_point_1_selection.selectedIndex].text;
	switch (assessment_point_1_selection_text){
	   case "Extreme applicability, alignment, or relevancy":
			final_premise_alignment_score += 8;
			break;
	   case "Significant applicability, alignment, or relevancy":
			final_premise_alignment_score += 6;
			break;
	   case "Moderate applicability, alignment, or relevancy": 
			final_premise_alignment_score += 4;
			break;
	   case "Low applicability, alignment, or relevancy": 
			final_premise_alignment_score += 2;
			break;
	   case "Not applicable, no alignment, or no relevancy": 
			final_premise_alignment_score += 0;
			break;
	}
	
	//Get 2nd result and convert to a score
	var assessment_point_2_selection = document.getElementById("premise_alignment_assessment_point_2");
	var assessment_point_2_selection_text = assessment_point_2_selection.options[assessment_point_2_selection.selectedIndex].text;
	switch (assessment_point_2_selection_text){
	   case "Extreme applicability, alignment, or relevancy":
			final_premise_alignment_score += 8;
			break;
	   case "Significant applicability, alignment, or relevancy":
			final_premise_alignment_score += 6;
			break;
	   case "Moderate applicability, alignment, or relevancy": 
			final_premise_alignment_score += 4;
			break;
	   case "Low applicability, alignment, or relevancy": 
			final_premise_alignment_score += 2;
			break;
	   case "Not applicable, no alignment, or no relevancy": 
			final_premise_alignment_score += 0;
			break;
	}
	
	//Get 3rd result and convert to a score
	var assessment_point_3_selection = document.getElementById("premise_alignment_assessment_point_3");
	var assessment_point_3_selection_text = assessment_point_3_selection.options[assessment_point_3_selection.selectedIndex].text;
	switch (assessment_point_3_selection_text){
	   case "Extreme applicability, alignment, or relevancy":
			final_premise_alignment_score += 8;
			break;
	   case "Significant applicability, alignment, or relevancy":
			final_premise_alignment_score +=6;
			break;
	   case "Moderate applicability, alignment, or relevancy": 
			final_premise_alignment_score += 4;
			break;
	   case "Low applicability, alignment, or relevancy": 
			final_premise_alignment_score += 2;
			break;
	   case "Not applicable, no alignment, or no relevancy": 
			final_premise_alignment_score += 0;
			break;
	}
	
	//Get 4th result and convert to a score
	var assessment_point_4_selection = document.getElementById("premise_alignment_assessment_point_4");
	var assessment_point_4_selection_text = assessment_point_4_selection.options[assessment_point_4_selection.selectedIndex].text;
	switch (assessment_point_4_selection_text){
	   case "Extreme applicability, alignment, or relevancy":
			final_premise_alignment_score += 8;
			break;
	   case "Significant applicability, alignment, or relevancy":
			final_premise_alignment_score += 6;
			break;
	   case "Moderate applicability, alignment, or relevancy": 
			final_premise_alignment_score += 4;
			break;
	   case "Low applicability, alignment, or relevancy": 
			final_premise_alignment_score += 2;
			break;
	   case "Not applicable, no alignment, or no relevancy": 
			final_premise_alignment_score += 0;
			break;
	}
	
	//Get 5th result and convert to a score
	var assessment_point_5_selection = document.getElementById("premise_alignment_assessment_point_5");
	var assessment_point_5_selection_text = assessment_point_5_selection.options[assessment_point_5_selection.selectedIndex].text;
	switch (assessment_point_5_selection_text){
	   case "Extreme applicability, alignment, or relevancy":
			final_premise_alignment_score += -8;
			break;
	   case "Significant applicability, alignment, or relevancy":
			final_premise_alignment_score += -6;
			break;
	   case "Moderate applicability, alignment, or relevancy": 
			final_premise_alignment_score += -4;
			break;
	   case "Low applicability, alignment, or relevancy": 
			final_premise_alignment_score += -2;
			break;
	   case "Not applicable, no alignment, or no relevancy": 
			final_premise_alignment_score += 0;
			break;
	}
	
	
		switch (true){
	   case (final_premise_alignment_score <= 10):
			var final_premise_alignment_score_text = "Low";
			break;
	   case (final_premise_alignment_score <= 18):
			var final_premise_alignment_score_text = "Medium";
			break;
	   case (final_premise_alignment_score <= 32):
			var final_premise_alignment_score_text = "High";
			break;
	}
	
	
	
	// Cue Selectors
	var final_cue_score = 0;
	var final_cue_score_text = "None"
	final_cue_score += document.getElementById("cue1").checked? 1 : 0;
	final_cue_score += document.getElementById("cue2").checked? 1 : 0;
	final_cue_score += document.getElementById("cue3").checked? 1 : 0;
	final_cue_score += document.getElementById("cue4").checked? 1 : 0;
	final_cue_score += document.getElementById("cue5").checked? 1 : 0;
	final_cue_score += document.getElementById("cue6").checked? 1 : 0;
	final_cue_score += document.getElementById("cue7").checked? 1 : 0;
	final_cue_score += document.getElementById("cue8").checked? 1 : 0;
	final_cue_score += document.getElementById("cue9").checked? 1 : 0;
	final_cue_score += document.getElementById("cue10").checked? 1 : 0;
	final_cue_score += document.getElementById("cue11").checked? 1 : 0;
	final_cue_score += document.getElementById("cue12").checked? 1 : 0;
	final_cue_score += document.getElementById("cue13").checked? 1 : 0;
	final_cue_score += document.getElementById("cue14").checked? 1 : 0;
	final_cue_score += document.getElementById("cue15").checked? 1 : 0;
	final_cue_score += document.getElementById("cue16").checked? 1 : 0;
	final_cue_score += document.getElementById("cue17").checked? 1 : 0;
	final_cue_score += document.getElementById("cue18").checked? 1 : 0;
	final_cue_score += document.getElementById("cue19").checked? 1 : 0;
	final_cue_score += document.getElementById("cue20").checked? 1 : 0;
	final_cue_score += document.getElementById("cue21").checked? 1 : 0;
	final_cue_score += document.getElementById("cue22").checked? 1 : 0;
	
	switch (true){
	   case (final_cue_score <= 0):
			var final_cue_score_text = "No Cues";
			break;		
	   case (final_cue_score <= 8):
			var final_cue_score_text = "Few";
			break;
	   case (final_cue_score <= 14):
			var final_cue_score_text = "Some";
			break;
	   case (final_cue_score <= 22):
			var final_cue_score_text = "Many";
			break;
	}
	
	//difficulty calulations
	var final_difficulty_score_text = "None"
	switch (true){
	   case (final_cue_score_text == "Few" && final_premise_alignment_score_text == "High"):
			final_difficulty_score_text = "Very Difficult"
			break;
	   case (final_cue_score_text == "Few" && final_premise_alignment_score_text == "Medium"):
			final_difficulty_score_text = "Very Difficult"
			break;			
	   case (final_cue_score_text == "Few" && final_premise_alignment_score_text == "Low"):
			final_difficulty_score_text = "Moderately Difficult"
			break;
	   case (final_cue_score_text == "Some" && final_premise_alignment_score_text == "High"):
			final_difficulty_score_text = "Very Difficult"
			break;
	   case (final_cue_score_text == "Some" && final_premise_alignment_score_text == "Medium"):
			final_difficulty_score_text = "Moderately Difficult"
			break;			
	   case (final_cue_score_text == "Some" && final_premise_alignment_score_text == "Low"):
			final_difficulty_score_text = "Moderately to Least Difficult"
			break;
	   case (final_cue_score_text == "Many" && final_premise_alignment_score_text == "High"):
			final_difficulty_score_text = "Moderately Difficult"
			break;
	   case (final_cue_score_text == "Many" && final_premise_alignment_score_text == "Medium"):
			final_difficulty_score_text = "Moderately Difficult"
			break;			
	   case (final_cue_score_text == "Many" && final_premise_alignment_score_text == "Low"):
			final_difficulty_score_text = "Least Difficult"
			break;
	}
	
	

	document.getElementById("result-premise-alignment").innerHTML = final_premise_alignment_score_text;
	document.getElementById("result-cue-count").innerHTML = final_cue_score_text;
	document.getElementById("result-difficulty").innerHTML = final_difficulty_score_text;

	
}



