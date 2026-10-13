import unittest
from datetime import date

import job_hunt


class JobHuntTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.config = job_hunt.load_config()
        cls.profile = job_hunt.load_profile()

    def test_canonical_urls_deduplicate_tracking_variants(self):
        first = "https://www.boards.greenhouse.io/careem/jobs/123?utm_source=mail&gh_jid=123#apply"
        second = "https://boards.greenhouse.io/careem/jobs/123?gh_jid=123"
        self.assertEqual(job_hunt.canonicalize_url(first), job_hunt.canonicalize_url(second))
        self.assertEqual(job_hunt.make_job_key("Careem", "Account Manager", "Dubai", first), job_hunt.make_job_key("Careem", "Account Manager", "Dubai", second))

    def test_fallback_deduplication_without_url(self):
        self.assertEqual(job_hunt.make_job_key("Careem", "Account Manager", "Dubai", None), job_hunt.make_job_key(" careem ", "ACCOUNT MANAGER", " DUBAI ", ""))

    def test_salary_monthly_annual_unknown_and_foreign(self):
        monthly = job_hunt.assess_salary("AED 6,500-8,000 per month", "", 6500)
        annual = job_hunt.assess_salary("AED 78,000-96,000 per annum", "", 6500)
        structured_annual = job_hunt.assess_salary({"min": 78000, "max": 96000, "currency": "AED", "interval": "year"}, "", 6500)
        under = job_hunt.assess_salary("AED 72,000 per year", "", 6500)
        ambiguous = job_hunt.assess_salary("AED 80,000", "", 6500)
        foreign = job_hunt.assess_salary("USD 90,000 per year", "", 6500)
        undisclosed = job_hunt.assess_salary(None, "Competitive salary", 6500)
        self.assertEqual(monthly["assessment"], "meets AED 6,500/month target")
        self.assertEqual(annual["monthly_min_aed"], 6500)
        self.assertEqual(structured_annual["monthly_min_aed"], 6500)
        self.assertTrue(under["exclude"])
        self.assertEqual(ambiguous["period"], "ambiguous")
        self.assertFalse(foreign["exclude"])
        self.assertIn("verification", foreign["assessment"])
        self.assertIn("undisclosed", undisclosed["assessment"])

    def test_salary_range_crossing_target_needs_review(self):
        salary = job_hunt.assess_salary("AED 6,000-8,000 per month", "", 6500)
        self.assertFalse(salary["exclude"])
        self.assertIn("crosses", salary["assessment"])

    def test_uae_only_and_city_priorities(self):
        self.assertEqual(job_hunt.classify_location("Dubai, United Arab Emirates"), ("Dubai", 1))
        self.assertEqual(job_hunt.classify_location("Sharjah, UAE"), ("Sharjah", 2))
        self.assertEqual(job_hunt.classify_location("Abu Dhabi, UAE"), ("Abu Dhabi", 3))
        self.assertEqual(job_hunt.classify_location("Ajman, UAE"), ("Ajman", 4))
        self.assertIsNone(job_hunt.classify_location("Bengaluru, India"))
        self.assertIsNone(job_hunt.classify_location("Remote"))

    def test_relevance_uses_title_evidence_and_years(self):
        job = {"title": "Product Operations Associate", "description": "3+ years in product operations. Own UAT, feature definition, workflows, dashboards, and Excel reporting."}
        score, reasons, exclusion = job_hunt.score_job(job, self.config, self.profile)
        self.assertFalse(exclusion)
        self.assertGreaterEqual(score, self.config["review_threshold"])
        self.assertTrue(any("documented evidence overlaps" in reason for reason in reasons))

    def test_unrelated_and_senior_roles_are_excluded(self):
        unrelated, _, reason = job_hunt.score_job({"title": "Software Engineer", "description": "Python backend development"}, self.config, self.profile)
        senior, _, senior_reason = job_hunt.score_job({"title": "Senior Product Operations Specialist", "description": "Product operations"}, self.config, self.profile)
        self.assertEqual(unrelated, 0)
        self.assertIn("outside", reason)
        self.assertEqual(senior, 0)
        self.assertIn("seniority", senior_reason)

    def test_expired_listing_exclusion(self):
        self.assertTrue(job_hunt.listing_is_expired({"application_deadline": "2025-01-01"}, date(2026, 10, 13)))
        self.assertFalse(job_hunt.listing_is_expired({"application_deadline": "2027-01-01"}, date(2026, 10, 13)))

    def test_upsert_preserves_status_and_date_found(self):
        old = {"job_key": "same", "company": "Careem", "role": "Account Manager", "location": "Dubai", "date_found": "2026-01-01", "status": "Applied", "next_action": "Follow up next week"}
        new = {**old, "date_found": "2026-10-13", "status": "New", "next_action": "Review", "match_score": 72}
        rows = job_hunt.upsert_tracker([old], [new])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["date_found"], "2026-01-01")
        self.assertEqual(rows[0]["status"], "Applied")
        self.assertEqual(rows[0]["next_action"], "Follow up next week")
        self.assertEqual(rows[0]["match_score"], 72)

    def test_city_priority_is_configurable_and_draft_keeps_experience_context(self):
        self.assertEqual(job_hunt.classify_location("Dubai, UAE", ["Abu Dhabi", "Dubai", "Other UAE"]), ("Dubai", 2))
        resume, outreach = job_hunt.build_drafts(
            {"role": "Product Operations Associate", "company": "Example Co"}, self.profile, self.config
        )
        self.assertIn("Accruon Consultant L.L.C.", resume)
        self.assertIn("Dec 2023 - Jan 2026", resume)
        self.assertIn("Cleantel Services", resume)
        self.assertIn("Draft for review only - not sent.", outreach)

    def test_review_draft_trigger_is_thresholded_and_not_repeated(self):
        self.assertTrue(job_hunt.should_generate_review_draft({"match_score": 55}, None, 55))
        self.assertFalse(job_hunt.should_generate_review_draft({"match_score": 54}, None, 55))
        previous = {"resume_draft": "resume.md", "outreach_draft": "outreach.md"}
        self.assertFalse(job_hunt.should_generate_review_draft({"match_score": 90}, previous, 55))

    def test_account_draft_prioritizes_client_evidence(self):
        resume, outreach = job_hunt.build_drafts(
            {"role": "Account Manager I", "company": "Example Co"}, self.profile, self.config
        )
        self.assertIn("SaaS onboarding", resume)
        self.assertIn("followed up with clients on gaps", outreach)
        self.assertIn("one contract renewal", outreach)
        self.assertNotIn("Built client master data", outreach)


if __name__ == "__main__":
    unittest.main()