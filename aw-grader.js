/*───────────────────────────────────────────────────────────────
  ArticuWrite — shared IELTS grader (aw-grader.js)

  The single source of truth for how an essay is graded. Both the student
  writing screen (write.html) and the reliability test page (test.html)
  call AW.grade(), so the test always measures the prompt that students
  actually get — a copy would silently drift out of sync.

  Depends on: aw-common.js (loaded first)
───────────────────────────────────────────────────────────────*/
(function (AW) {
  'use strict';

  function normaliseScores(r){
    if(r.scores){
      var avg=(r.scores.TR+r.scores.CC+r.scores.LR+r.scores.GRA)/4;
      // IELTS convention: .25–.74 → .5, otherwise nearest whole band
      var whole=Math.floor(avg), frac=avg-whole;
      r.overall = frac<0.25 ? whole : (frac<0.75 ? whole+0.5 : whole+1);
    }
    return r;
  }
  async function gradeWithGemini(payload, geminiKey){
    var text=payload.writing||'', question=payload.question||'';
    var systemInstruction='You are a strict IELTS Writing examiner. Grade only the essay provided by the student. Every piece of feedback must quote directly from that essay. Never invent examples or use content from other essays.\n\n'+
      'IELTS BAND DESCRIPTORS:\n'+
      'TR/TA: Band 7=covers all parts, clear position, ideas developed. Band 6=addresses parts but some inadequately. Band 5=partial, limited. Band 4=minimal.\n'+
      'CC: Band 7=logical, varied cohesion, good paragraphing. Band 6=mostly coherent. Band 5=some org. Band 4=incoherent.\n'+
      'LR: Band 7=flexible range, minor errors. Band 6=adequate. Band 5=limited. Band 4=basic.\n'+
      'GRA: Band 7=complex structures, frequent error-free. Band 6=mix. Band 5=frequent errors. Band 4=many errors.\n\n'+
      'SCORING RULES:\n- Components: whole numbers or X.5 if borderline.\n- Overall = average of 4, round to 0.5.\n- Most students score 5-6. Band 7 = genuinely good. Band 8+ very rare.\n- Grammar errors → GRA ≤ 6. Basic vocab → LR ≤ 6.\n\n'+
      '[CONDITIONAL RULES — enforce strictly, do not overlook any]\n'+
      '1) LEXICAL RESOURCE & GRAMMAR — HARD LIMITS:\n'+
      '- Spelling: count the spelling mistakes. IF more than 5 (i.e. frequent, causing difficulty for the reader) THEN cap LR at Band 5 and warn in Vietnamese: "Sai chính tả trên 5 lỗi không thể đạt Band 6 tiêu chí Vocab". IF 1-5 (occasional slips) THEN cap LR at Band 6 and list every error found.\n'+
      '- IF there are basic punctuation errors or failure to capitalize proper nouns (e.g. "zoom" instead of "Zoom", "google meet" instead of "Google Meet") THEN cap GRA at Band 6.\n'+
      '- IF the student uses contractions (it\'s, don\'t) or informal vocabulary (e.g. "costly mistakes", "thanks to" instead of "due to") THEN penalize and require formal academic alternatives.\n'+
      '- IF a body paragraph consists entirely of simple sentences without subordinating conjunctions (although, whereas, provided that) or relative clauses (which, who, that) THEN flag as "Lack of Grammatical Variety" and warn that Band 7 in GRA requires a mix of complex sentence structures.\n'+
      '- Repetition: IF a topic-specific keyword (e.g. "online meetings") appears more than 2 times in one body paragraph, OR more than once in the Intro/Conclusion, THEN flag it as a repetition error and suggest pronouns or contextual synonyms — BUT only when the repetition shows a lack of vocabulary. ACCEPT the repetition if replacing it with a synonym would distort the meaning (e.g. fixed technical terms).\n'+
      '- IF the word "nowadays" is used THEN flag it as a cliché and demand removal or replacement.\n'+
      '- IF the student uses an unnatural word combination (e.g. "do a mistake" instead of "make a mistake", "heavy traffic jam" instead of "heavy traffic") THEN quote the phrase, label it "Unnatural Collocation", and give the correct academic alternative.\n'+
      '- IF the student uses informal phrasal verbs ("look into", "come up with", "get rid of") in an academic context THEN penalize the tone and suggest formal single-word equivalents ("investigate", "develop/invent", "eliminate").\n'+
      '2) TASK RESPONSE (CONTENT & LOGIC):\n'+
      '- IF an example or argument shifts context away from the prompt (prompt about "business" but student writes about "education"/"students") THEN flag as "Sai bối cảnh / Off-topic".\n'+
      '- IF an argument lacks real-world logic (e.g. "traffic noise" distracting someone "working from home", or "absent-mindedness" as a professional excuse) THEN counter-argue with a critical question exposing the logical flaw.\n'+
      '- IF the student introduces a new idea (especially at the end of a paragraph) without a supporting sentence THEN flag as "Lỗi liệt kê" (Listing error): the idea lacks development.\n'+
      '- IF the prompt does NOT ask to discuss "benefits and drawbacks" BUT the student includes them in the introductory paraphrase THEN flag as an inaccurate paraphrase.\n'+
      '- IF the opinion in the Conclusion contradicts or shifts away from the thesis statement in the Introduction THEN flag as "Inconsistent Position" and cap Task Response at Band 6.\n'+
      '3) COHERENCE & COHESION:\n'+
      '- IF an example merely rewrites the previous sentence with different vocabulary without adding specific concrete details THEN flag as "Invalid Example": an example must be more specific than the claim it supports.\n'+
      '- IF a Topic Sentence lists the specific main ideas to be discussed THEN advise keeping it broader with clear directional words (e.g. skeptical, advocate).\n'+
      '- IF a sentence states a result ("As a result, businesses reduce expenses") BUT the preceding sentence did not mention the corresponding cause (money/costs) THEN flag as a logical disconnect.\n'+
      '- IF transition words are combined redundantly ("Consequently, this...") THEN correct to a single cohesive device.\n'+
      '- Conclusion: IF it introduces new ideas THEN penalize. IF it copies body sentences word-for-word THEN advise paraphrasing. Concisely summarizing the main points IS acceptable and must NOT be penalized.\n'+
      '- IF the student starts almost every sentence with a basic transition ("Firstly,", "Secondly,", "Moreover,", "Furthermore,") THEN flag as "Mechanical Cohesion" and suggest referencing pronouns (this, these, such) or advanced cohesive structures ("Not only..., but...", "Another compelling reason is...").\n'+
      '- IF a body paragraph has fewer than 3 sentences OR discusses two completely unrelated main ideas THEN flag as "Paragraphing Error" and advise ONE central idea per paragraph with sufficient development.\n'+
      '[FEEDBACK FORMAT]\n'+
      'When any rule triggers, quote the exact sentence containing the error, use an arrow "=>", and give a precise, constructive correction based ONLY on the rules above. Keep the tone professional, direct and academic. Apply caps to the numeric scores accordingly — caps set the MAXIMUM; genuine weaknesses may score lower.\n\n'+
      'Return ONLY a JSON object with fields: scores(TR,CC,LR,GRA), overall, band_description, overall_feedback_vi, tr_comments(array of {paragraph_role,assessment_vi,suggestion_en,quote}), gra_errors(array of {wrong,correct,explanation_vi}), lr_issues(array of {original,better,explanation_vi,alternatives}), cc_feedback({assessment_vi,issues,suggestions}), corrected_text, repeated_errors_vi.\n'+
      'TOKEN RULES for tr_comments — follow strictly to keep output short:\n'+
      '- Do NOT quote whole paragraphs. For each comment set "paragraph_role" to the paragraph label ONLY: "Opening", "Body 1", "Body 2", "Body 3"... , or "Conclusion". Standard mapping: a 4-paragraph essay = Opening, Body 1, Body 2, Conclusion; a 5-paragraph essay = Opening, Body 1, Body 2, Body 3, Conclusion.\n'+
      '- Leave "quote" EMPTY unless you must point to ONE specific problematic sentence — then put ONLY that single sentence in "quote" (never the whole paragraph).\n'+
      '- "assessment_vi" is your comment on that paragraph. Keep it concise.';
    // Strip HTML tags from question (may contain <b>,<i> from rich text editor)
    var tmp=document.createElement('div'); tmp.innerHTML=question;
    var cleanQ=(tmp.textContent||tmp.innerText||question).replace(/\s+/g,' ').trim();
    var attemptNote = payload.attempt >= 2
      ? '\n\nThis is revision attempt '+payload.attempt+'. In "repeated_errors_vi" (Vietnamese), note whether the student repeated the same TYPES of mistakes as a typical earlier draft (grammar/vocab/coherence) and encourage improvement. For attempt 1, leave repeated_errors_vi empty.'
      : '\n\nThis is the first attempt. Leave repeated_errors_vi empty. Give COMPLETE feedback on all 4 criteria so the student learns from the start.';
    var userMessage='Grade this IELTS essay.\n\nTASK PROMPT: '+cleanQ+'\n\nSTUDENT ESSAY TO GRADE:\n---\n'+text+'\n---\n\nAll feedback must reference the student essay above and be thorough. Task Response score MUST assess how directly the essay addresses the given task prompt. If the essay goes off-topic or misses the task, TR ≤ 5.'+attemptNote;
    // Teacher's context rules override default IELTS strictness
    if (payload.aiNotes && payload.aiNotes.trim()) {
      userMessage += '\n\n=== TEACHER\'S GRADING RULES (HIGHEST PRIORITY — these OVERRIDE the default IELTS strictness above) ===\n'+
        payload.aiNotes.trim()+
        '\n\nYou MUST follow these teacher rules. Do NOT flag, mark, or deduct points for anything the teacher has explicitly allowed or told you to ignore. Only report errors that remain genuine problems given these rules. This keeps feedback appropriate for the class context and avoids overwhelming the student with irrelevant corrections.';
    }
    var prompt=systemInstruction+'\n\n'+userMessage;
    // ~30-token CEFR vocabulary summary (measured client-side) for reference
    if (payload.vocabSummary && payload.vocabSummary.trim()) {
      prompt += '\n\n=== VOCABULARY DATA (for Lexical Resource reference) ===\n' + payload.vocabSummary.trim();
    }

    // Grading must be as repeatable as possible, so decode near-greedily:
    // temperature alone only sharpens the distribution — topK/topP are what
    // actually stop the model sampling alternative wordings/scores.
    // (Note: even at these settings the API is not bit-for-bit deterministic.)
    var resp=await fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key='+geminiKey,{
      method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({
        contents:[{parts:[{text:prompt}]}],
        generationConfig:{ temperature:0.1, topK:1, topP:0.1, candidateCount:1, maxOutputTokens:12000 }
      })
    });
    if(!resp.ok){var err=await resp.json().catch(function(){return{};});throw new Error('Gemini error: '+((err.error&&err.error.message)||resp.status));}
    var data=await resp.json();
    var raw=(((data.candidates||[])[0]||{}).content||{}).parts;
    raw=raw&&raw[0]?raw[0].text:'';
    if(!raw) throw new Error('Gemini returned empty. Try again.');
    var clean=raw.replace(/<thinking>[\s\S]*?<\/thinking>/gi,'').replace(/```json\s*/gi,'').replace(/```\s*/g,'').trim();
    var si=clean.indexOf('{'),depth=0,ei=-1;
    for(var ci=si;ci<clean.length;ci++){if(clean[ci]==='{')depth++;else if(clean[ci]==='}'){depth--;if(depth===0){ei=ci;break;}}}
    if(si===-1||ei===-1) throw new Error('Could not parse JSON from Gemini.');
    return normaliseScores(JSON.parse(clean.substring(si,ei+1)));
  }
  // Public API
  AW.grade = gradeWithGemini;
  AW.gradeNormalise = normaliseScores;

})(window.AW = window.AW || {});
