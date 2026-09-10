import unittest

from cis_benchmark_converter import (
    extract_policy_name_and_recommended_setting,
    extract_recommendations,
)


class TestTitleExtraction(unittest.TestCase):
    def test_extract_policy_and_setting_from_ensure_title(self):
        title = "Ensure 'Enforce password history' is set to '24 or more password(s)' (Automated)"
        policy, setting = extract_policy_name_and_recommended_setting(title)
        self.assertEqual(policy, "Enforce password history")
        self.assertEqual(setting, "24 or more password(s)")

    def test_extract_policy_and_setting_from_configure_title(self):
        title = "Configure 'Accounts: Rename guest account' (Automated)"
        policy, setting = extract_policy_name_and_recommended_setting(title)
        self.assertEqual(policy, "Accounts: Rename guest account")
        self.assertEqual(setting, "")

    def test_extract_policy_and_setting_keeps_dc_only(self):
        title = "Ensure 'Domain controller: LDAP server signing requirements' is set to 'Require signing' (DC Only) (Automated)"
        policy, setting = extract_policy_name_and_recommended_setting(title)
        self.assertEqual(policy, "Domain controller: LDAP server signing requirements (DC Only)")
        self.assertEqual(setting, "Require signing")

    def test_extract_recommendations_populates_new_fields(self):
        full_text = "\n".join(
            [
                "1.2.3 (L1) Ensure 'Enforce password history' is set to '24 or more password(s)' (Automated)",
                "Profile Applicability:",
                "Description:",
                "Sample",
            ]
        )

        recommendations = extract_recommendations(full_text)
        self.assertEqual(len(recommendations), 1)
        self.assertEqual(recommendations[0]["Policy Name"], "Enforce password history")
        self.assertEqual(recommendations[0]["Recommended Setting"], "24 or more password(s)")


if __name__ == "__main__":
    unittest.main()
