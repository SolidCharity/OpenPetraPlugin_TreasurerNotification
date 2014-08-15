SELECT PUB_pm_staff_data.*
FROM PUB_pm_staff_data, PUB_p_person
WHERE PUB_p_person.p_family_key_n = ?
   AND PUB_pm_staff_data.p_partner_key_n = PUB_p_person.p_partner_key_n
