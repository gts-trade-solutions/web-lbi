// Standard, fixed report wording for RACE LBI route-feasibility reports.
//
// The survey company reuses the SAME certification / conclusion text for every
// job and only changes the route's origin ("From") and destination ("To").
// So the Conclusion is generated from this fixed template with just {FROM} and
// {TO} substituted in — nothing else is editable. The Objective keeps the
// standard sentence as a starting point (the cargo dimensions are filled in per
// job, since a feasibility objective is meaningless without the cargo size).
//
// Captured from the RACE "… to Mundra Port" detail report (Conclusion &
// Certification pages), generalised to a single From → To route.

/** Default Objective sentence — pre-filled; the L/W/H/WT are edited per job. */
export const STANDARD_OBJECTIVE_DEFAULT =
  "To study and evaluate the feasibility of moving cargo measuring dimensions of Consignment: L __ m X W __ m X H __ m including trailer - WT __ Tons";

function escapeHtml(s: string): string {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * Build the fixed standard "Conclusion & Certification" HTML for a route,
 * substituting only the origin (from) and destination (to). Everything else is
 * the standard RACE wording and is not meant to be edited.
 */
export function buildStandardConclusionHtml(from: string, to: string): string {
  const FROM = escapeHtml((from || "").trim()) || "[Origin]";
  const TO = escapeHtml((to || "").trim()) || "[Destination]";

  return [
    `<p><strong>CONCLUSION: ROUTE FEASIBILITY SURVEY – SUMMARY REPORT</strong></p>`,
    `<p><strong>Project:</strong> ${FROM} to ${TO}</p>`,
    `<p><strong>Survey Certification &amp; Scope:</strong> This is to certify that a comprehensive route feasibility survey has been conducted by the RACE Survey Team for the proposed transportation of project cargo from ${FROM} to ${TO}.</p>`,
    `<p>The survey covered detailed on-ground inspections, including:</p>`,
    `<ul>`,
    `<li>Route geometry and carriageway assessment</li>`,
    `<li>Horizontal and vertical clearance verification</li>`,
    `<li>Bridge and culvert condition assessment</li>`,
    `<li>Turning radius and gradient analysis</li>`,
    `<li>Identification of critical bottlenecks and risk points / bridge underpasses</li>`,
    `<li>Overall movement suitability for Over Dimensional Cargo (ODC)</li>`,
    `</ul>`,
    `<p>Based on the route feasibility assessments conducted for the transportation of Over Dimensional Cargo (ODC) from ${FROM} to ${TO}, the following key observations and recommendations have been identified. The assessment identified certain infrastructure constraints along the proposed transportation route that require appropriate modifications and mitigation measures to facilitate safe cargo movement.</p>`,
    `<ul>`,
    `<li><strong>Factory / Exit Constraints:</strong> Due to the overall cargo and vehicle combination length, civil modifications are recommended at the exit to ensure safe manoeuvrability.</li>`,
    `<li><strong>Underpass Clearance:</strong> Height restrictions at certain underpasses necessitate suitable bypass arrangements, subject to case-specific feasibility assessments.</li>`,
    `<li><strong>Pedestrian Foot Overbridges:</strong> Clearance constraints at identified locations may require temporary removal, alternative bypass arrangements, or road-level modifications, including excavation, depending on site conditions.</li>`,
    `<li><strong>Toll Plaza Constraints:</strong> Temporary modifications to toll plaza infrastructure are recommended wherever necessary to accommodate the overall transportation width.</li>`,
    `<li><strong>Other Infrastructure Obstacles:</strong> Minor clearance works, including temporary relocation or removal of signboards, adjustment of low-tension (LT) electrical lines, and other roadside obstructions, may be required to ensure unobstructed movement.</li>`,
    `</ul>`,
    `<p><strong>Overall Route Feasibility Conclusion</strong></p>`,
    `<p>Based on the route assessments conducted, transportation of the proposed cargo from ${FROM} to ${TO} is considered technically feasible using the proposed hydraulic trailer and vehicle combinations specified in the respective General Arrangement Drawings (GADs).</p>`,
    `<p>However, the feasibility is subject to the successful implementation of the recommended civil works, infrastructure modifications, bypass arrangements, and other site-specific mitigation measures identified during the route surveys.</p>`,
    `<p>All necessary statutory approvals, permissions from the relevant authorities, utility clearances, and traffic management arrangements must be obtained and completed before commencement of transportation.</p>`,
    `<p>The final movement shall be subject to verification of the actual cargo dimensions, gross transportation weight, hydraulic trailer configuration, and prevailing route conditions prior to execution.</p>`,
    `<p><strong>Issued By</strong><br />For and on behalf of:<br />RACE LBI – Route Logistics &amp; Engineering Division<br />Website: www.raceinnovations.in</p>`,
  ].join("");
}
