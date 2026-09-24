// Common Core ELA standards catalog and where each is taught.
// Edit STANDARD_MAP to change which lessons teach which standard: map[child][subject key][week] = [codes].
// Sub-parts (e.g. L.4.1a) roll up to their parent standard. Speaking & Listening (SL) is listed but has no lessons mapped: it isn't a subject in the current outline.
const STANDARD_GRADE = { kenley: 7, adelyn: 4 };
const STANDARD_CATALOG = {
 "4": [
  {
   "id": "RL",
   "name": "Reading: Literature",
   "codes": [
    [
     "RL.4.1",
     "Refer to details and examples in a story when explaining it and drawing inferences"
    ],
    [
     "RL.4.2",
     "Determine a theme from details; summarize the story"
    ],
    [
     "RL.4.3",
     "Describe a character, setting, or event in depth using specific details"
    ],
    [
     "RL.4.4",
     "Determine word and phrase meanings in a text, including allusions to mythology"
    ],
    [
     "RL.4.5",
     "Explain how poems, plays, and prose differ, using structural terms"
    ],
    [
     "RL.4.6",
     "Compare first-person and third-person points of view"
    ],
    [
     "RL.4.7",
     "Connect a story to a visual or oral version of it"
    ],
    [
     "RL.4.9",
     "Compare and contrast themes and events in stories from different cultures"
    ],
    [
     "RL.4.10",
     "Read and understand grade-level literature independently and proficiently"
    ]
   ]
  },
  {
   "id": "RI",
   "name": "Reading: Informational Text",
   "codes": [
    [
     "RI.4.1",
     "Refer to details and examples in informational text and draw inferences"
    ],
    [
     "RI.4.2",
     "Determine the main idea, explain how details support it, and summarize"
    ],
    [
     "RI.4.3",
     "Explain events, ideas, or steps in a text, including what happened and why"
    ],
    [
     "RI.4.4",
     "Determine the meaning of academic and subject-specific words"
    ],
    [
     "RI.4.5",
     "Describe a text’s structure (chronology, comparison, cause/effect, problem/solution)"
    ],
    [
     "RI.4.6",
     "Compare a firsthand and a secondhand account of the same topic"
    ],
    [
     "RI.4.7",
     "Interpret charts, graphs, diagrams, and timelines and explain how they help"
    ],
    [
     "RI.4.8",
     "Explain how an author uses reasons and evidence to support points"
    ],
    [
     "RI.4.9",
     "Combine information from two texts on the same topic"
    ],
    [
     "RI.4.10",
     "Read and understand grade-level informational text independently"
    ]
   ]
  },
  {
   "id": "RF",
   "name": "Reading: Foundational Skills",
   "codes": [
    [
     "RF.4.3",
     "Use phonics and word-analysis skills to decode multisyllable words"
    ],
    [
     "RF.4.4",
     "Read with enough accuracy and fluency to support understanding"
    ]
   ]
  },
  {
   "id": "W",
   "name": "Writing",
   "codes": [
    [
     "W.4.1",
     "Write opinion pieces with a point of view supported by reasons"
    ],
    [
     "W.4.2",
     "Write informative pieces that explain a topic clearly"
    ],
    [
     "W.4.3",
     "Write narratives with descriptive details and clear sequence"
    ],
    [
     "W.4.4",
     "Produce clear, organized writing that fits the task and audience"
    ],
    [
     "W.4.5",
     "With guidance, plan, revise, and edit writing"
    ],
    [
     "W.4.6",
     "Use technology to produce and publish writing (including keyboarding)"
    ],
    [
     "W.4.7",
     "Do short research projects that build knowledge"
    ],
    [
     "W.4.8",
     "Gather information from sources, take notes, and sort evidence"
    ],
    [
     "W.4.9",
     "Use evidence from texts to support analysis and reflection"
    ],
    [
     "W.4.10",
     "Write regularly over short and extended time frames"
    ]
   ]
  },
  {
   "id": "L",
   "name": "Language (grammar, spelling, vocabulary)",
   "codes": [
    [
     "L.4.1",
     "Use correct grammar and usage (pronouns, tenses, prepositional phrases, complete sentences)"
    ],
    [
     "L.4.2",
     "Use correct capitalization, punctuation, and spelling"
    ],
    [
     "L.4.3",
     "Choose words and sentence styles for precision and effect"
    ],
    [
     "L.4.4",
     "Figure out unknown words using context, Greek and Latin roots and affixes, and references"
    ],
    [
     "L.4.5",
     "Understand figurative language, word relationships, and shades of meaning"
    ],
    [
     "L.4.6",
     "Learn and use grade-level academic and subject-specific vocabulary"
    ]
   ]
  },
  {
   "id": "SL",
   "name": "Speaking & Listening",
   "codes": [
    [
     "SL.4.1",
     "Take part in collaborative discussions: prepare, follow rules, ask and answer questions, build on others’ ideas"
    ],
    [
     "SL.4.2",
     "Paraphrase a text read aloud or information presented in other media"
    ],
    [
     "SL.4.3",
     "Identify the reasons and evidence a speaker gives to support points"
    ],
    [
     "SL.4.4",
     "Report on a topic or tell a story with facts and details, speaking clearly at an understandable pace"
    ],
    [
     "SL.4.5",
     "Add audio and visual displays to presentations when they help"
    ],
    [
     "SL.4.6",
     "Tell when to use formal English and when informal speech is fine"
    ]
   ],
   "manual": true,
   "note": "Speaking & Listening is not a subject in the current outline, so these are not mapped to any lesson. Address them manually (discussions, presentations, read-alouds) or develop lessons for them later."
  }
 ],
 "7": [
  {
   "id": "RL",
   "name": "Reading: Literature",
   "codes": [
    [
     "RL.7.1",
     "Cite several pieces of textual evidence to support analysis and inferences"
    ],
    [
     "RL.7.2",
     "Determine a theme or central idea, trace its development, and summarize"
    ],
    [
     "RL.7.3",
     "Analyze how story elements (setting, plot, characters) interact"
    ],
    [
     "RL.7.4",
     "Determine word meanings, including figurative and connotative; analyze sound devices"
    ],
    [
     "RL.7.5",
     "Analyze how a poem’s or drama’s form and structure add to its meaning"
    ],
    [
     "RL.7.6",
     "Analyze how an author develops and contrasts different points of view"
    ],
    [
     "RL.7.7",
     "Compare a written story with its audio, film, or staged version"
    ],
    [
     "RL.7.9",
     "Compare fiction with a historical account of the same time, place, or character"
    ],
    [
     "RL.7.10",
     "Read and understand grades 6–8 literature independently and proficiently"
    ]
   ]
  },
  {
   "id": "RI",
   "name": "Reading: Informational Text",
   "codes": [
    [
     "RI.7.1",
     "Cite several pieces of textual evidence to support analysis and inferences"
    ],
    [
     "RI.7.2",
     "Determine two or more central ideas, trace their development, and summarize objectively"
    ],
    [
     "RI.7.3",
     "Analyze how individuals, events, and ideas interact in a text"
    ],
    [
     "RI.7.4",
     "Determine word meanings (figurative, connotative, technical) and the effect of word choice"
    ],
    [
     "RI.7.5",
     "Analyze how a text is organized and how sections build the whole"
    ],
    [
     "RI.7.6",
     "Determine an author’s point of view or purpose and how they set it apart from others"
    ],
    [
     "RI.7.7",
     "Compare a text with its audio, video, or multimedia version"
    ],
    [
     "RI.7.8",
     "Trace and evaluate an argument, judging whether reasoning and evidence hold up"
    ],
    [
     "RI.7.9",
     "Analyze how two authors on the same topic present key information differently"
    ],
    [
     "RI.7.10",
     "Read and understand grades 6–8 informational text independently"
    ]
   ]
  },
  {
   "id": "W",
   "name": "Writing",
   "codes": [
    [
     "W.7.1",
     "Write arguments that support claims with clear reasons and relevant evidence"
    ],
    [
     "W.7.2",
     "Write informative texts that organize and analyze ideas clearly"
    ],
    [
     "W.7.3",
     "Write narratives with effective technique, details, and sequence"
    ],
    [
     "W.7.4",
     "Produce clear, coherent writing suited to task, purpose, and audience"
    ],
    [
     "W.7.5",
     "Plan, revise, edit, and rewrite to strengthen writing"
    ],
    [
     "W.7.6",
     "Use technology to produce, publish, and cite sources"
    ],
    [
     "W.7.7",
     "Do short research projects using several sources"
    ],
    [
     "W.7.8",
     "Gather and assess sources; quote or paraphrase without plagiarizing"
    ],
    [
     "W.7.9",
     "Use evidence from texts to support analysis and research"
    ],
    [
     "W.7.10",
     "Write regularly over short and extended time frames"
    ]
   ]
  },
  {
   "id": "L",
   "name": "Language (grammar, spelling, vocabulary)",
   "codes": [
    [
     "L.7.1",
     "Use correct grammar: phrases and clauses, sentence types, and modifier placement"
    ],
    [
     "L.7.2",
     "Use correct capitalization, punctuation, and spelling"
    ],
    [
     "L.7.3",
     "Use precise, concise language and vary sentence style"
    ],
    [
     "L.7.4",
     "Figure out unknown words using context, Greek and Latin roots and affixes, and references"
    ],
    [
     "L.7.5",
     "Understand figurative language, word relationships, and shades of meaning"
    ],
    [
     "L.7.6",
     "Learn and use grade-level academic and subject-specific vocabulary"
    ]
   ]
  },
  {
   "id": "SL",
   "name": "Speaking & Listening",
   "codes": [
    [
     "SL.7.1",
     "Take part in collaborative discussions: prepare, follow rules, pose questions, respond to new information"
    ],
    [
     "SL.7.2",
     "Analyze main ideas and supporting details in diverse media and explain how they clarify a topic"
    ],
    [
     "SL.7.3",
     "Outline a speaker’s argument and judge whether the reasoning and evidence are sound"
    ],
    [
     "SL.7.4",
     "Present claims and findings in a focused, coherent way with clear speaking (eye contact, volume, pronunciation)"
    ],
    [
     "SL.7.5",
     "Use multimedia and visual displays to clarify claims and findings"
    ],
    [
     "SL.7.6",
     "Adapt speech to different contexts and tasks, using formal English when appropriate"
    ]
   ],
   "manual": true,
   "note": "Speaking & Listening is not a subject in the current outline, so these are not mapped to any lesson. Address them manually (discussions, presentations, read-alouds) or develop lessons for them later."
  }
 ]
};
const STANDARD_MAP = {"kenley":{"vocab":{"1":["L.7.4","L.7.6"],"2":["L.7.4","L.7.6"],"3":["L.7.4","L.7.6"],"4":["L.7.4","L.7.6"],"5":["L.7.4","L.7.6"]},"spelling":{"1":["L.7.2"],"2":["L.7.2"],"3":["L.7.2"],"4":["L.7.2"],"5":["L.7.2"]},"grammar":{"1":["L.7.1"],"2":["L.7.1"],"3":["L.7.1"],"4":["L.7.1"],"5":["L.7.1"]},"reading":{"1":["RL.7.1","RL.7.3","RL.7.6"],"2":["RL.7.1","RL.7.3","RL.7.6"],"3":["RL.7.1","RL.7.3","RL.7.6"],"4":["RL.7.1","RL.7.3","RL.7.6"],"5":["RL.7.1","RL.7.2","RL.7.3","RL.7.6","RL.7.10"]},"writing":{"1":["W.7.1","W.7.4"],"2":["W.7.1","W.7.4","W.7.10"],"3":["W.7.1","W.7.4"],"4":["W.7.5","W.7.4"],"5":["W.7.1","W.7.4","W.7.5","W.7.10"]}},"adelyn":{"vocab":{"1":["L.4.4","L.4.6"],"2":["L.4.4","L.4.6"],"3":["L.4.4","L.4.6"],"4":["L.4.4","L.4.6"],"5":["L.4.4","L.4.6"],"6":["L.4.4","L.4.6"],"7":["L.4.4","L.4.6"],"8":["L.4.4","L.4.6"],"9":["L.4.4","L.4.6"],"10":["L.4.4","L.4.6"],"11":["L.4.4","L.4.6"],"12":["L.4.4","L.4.6"],"13":["L.4.4","L.4.6"],"14":["L.4.4","L.4.6"],"15":["L.4.4","L.4.6"],"16":["L.4.4","L.4.6"],"17":["L.4.4","L.4.6"],"18":["L.4.4","L.4.6"],"19":["L.4.4","L.4.6"],"20":["L.4.4","L.4.6"],"21":["L.4.4","L.4.6"],"22":["L.4.4","L.4.6"],"23":["L.4.4","L.4.6"],"24":["L.4.4","L.4.6"],"25":["L.4.4","L.4.6"],"26":["L.4.4","L.4.6"],"27":["L.4.4","L.4.6"],"28":["L.4.4","L.4.6"],"29":["L.4.4","L.4.6"],"30":["L.4.4","L.4.6"],"31":["L.4.4","L.4.6"],"32":["L.4.4","L.4.6"],"33":["L.4.4","L.4.6"],"34":["L.4.4","L.4.6"],"35":["L.4.4","L.4.6"],"36":["L.4.4","L.4.6"]},"spelling":{"1":["L.4.2","RF.4.3"],"2":["L.4.2","RF.4.3"],"3":["L.4.2","RF.4.3"],"4":["L.4.2","RF.4.3"],"5":["L.4.2","RF.4.3"],"6":["L.4.2","RF.4.3"],"7":["L.4.2","RF.4.3"],"8":["L.4.2","RF.4.3"],"9":["L.4.2","RF.4.3"],"10":["L.4.2","RF.4.3"],"11":["L.4.2","RF.4.3"],"12":["L.4.2","RF.4.3"],"13":["L.4.2","RF.4.3"],"14":["L.4.2","RF.4.3"],"15":["L.4.2","RF.4.3"],"16":["L.4.2","RF.4.3"],"17":["L.4.2","RF.4.3"],"18":["L.4.2","RF.4.3"],"19":["L.4.2","RF.4.3"],"20":["L.4.2","RF.4.3"],"21":["L.4.2","RF.4.3"],"22":["L.4.2","RF.4.3"],"23":["L.4.2","RF.4.3"],"24":["L.4.2","RF.4.3"],"25":["L.4.2","RF.4.3"],"26":["L.4.2","RF.4.3"],"27":["L.4.2","RF.4.3"]},"grammar":{"1":["L.4.6","L.4.1"],"2":["L.4.1"],"3":["L.4.1"],"4":["L.4.1","W.4.3"],"5":["L.4.1","W.4.3"],"6":["L.4.1","W.4.3"],"7":["L.4.1"],"8":["L.4.1"],"9":["L.4.1"],"10":["L.4.1","L.4.6"]},"reading":{"1":["RL.4.3","RF.4.3","RF.4.4"],"2":["RL.4.3","RF.4.3","RF.4.4"],"3":["RL.4.3","RF.4.3","RF.4.4"],"4":["RL.4.7","RF.4.3","RF.4.4"],"5":["RL.4.5","RF.4.3","RF.4.4"],"6":["RL.4.5","RF.4.3","RF.4.4"],"7":["RL.4.7","RF.4.3","RF.4.4"],"8":["RL.4.3","RF.4.3","RF.4.4"],"9":["RI.4.1","RI.4.2","RF.4.3","RF.4.4"],"10":["RI.4.2","RF.4.3","RF.4.4"],"11":["RI.4.4","RF.4.3","RF.4.4"],"12":["RL.4.6","RF.4.3","RF.4.4"],"13":["RL.4.6","RF.4.3","RF.4.4"],"14":["RL.4.6","RF.4.3","RF.4.4"],"15":["RL.4.5","RF.4.3","RF.4.4"],"16":["RL.4.5","RF.4.3","RF.4.4"],"17":["RI.4.5","RF.4.3","RF.4.4"],"18":["RI.4.3","RF.4.3","RF.4.4"],"19":["RI.4.6","RF.4.3","RF.4.4"],"20":["RL.4.4","RF.4.3","RF.4.4"],"21":["RL.4.1","RL.4.3","RF.4.3","RF.4.4"],"22":["RL.4.2","RF.4.3","RF.4.4"],"23":["RL.4.2","RF.4.3","RF.4.4"],"24":["RL.4.4","RF.4.3","RF.4.4"],"25":["RL.4.3","RL.4.9","RF.4.3","RF.4.4"],"26":["RL.4.9","RF.4.3","RF.4.4"],"27":["RL.4.9","RF.4.3","RF.4.4"],"28":["RI.4.9","RF.4.3","RF.4.4"],"29":["RI.4.7","RF.4.3","RF.4.4"],"30":["RL.4.10","RF.4.3","RF.4.4"],"31":["RL.4.3","RL.4.6","RF.4.3","RF.4.4"],"32":["RL.4.5","RF.4.3","RF.4.4"],"33":["RL.4.1","RL.4.10","RF.4.3","RF.4.4"],"34":["RL.4.3","RL.4.5","RF.4.3","RF.4.4"],"35":["RL.4.9","RF.4.3","RF.4.4"],"36":["RF.4.4","RF.4.3"]},"writing":{"1":["W.4.3","L.4.3","W.4.10"],"2":["W.4.3","L.4.3","L.4.1"],"3":["W.4.3","L.4.1","L.4.3"],"4":["W.4.3","W.4.4","L.4.3"]}}};
