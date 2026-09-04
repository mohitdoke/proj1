-- ============================================================
-- Correct fund membership using the user-provided "Fund details.xlsx"
-- (the authoritative source, not the earlier default/guessed mapping).
--
-- Real mapping:
--   Fund 1: Leegality, FinBox, EasyRewardz, Multipl, Fastsurance,
--           Riskcovry, Castler, GrayQuest, Vitra.ai
--           (+ Finvu, Data Sutram - not yet in `companies`, added separately
--            once their data/company rows exist)
--   Fund 2: Leegality, GrayQuest, Fundamento, Knight FinTech, Traqcheck
--           (+ Finvu - not yet in `companies`)
--
-- i.e. Leegality and GrayQuest belong to BOTH funds; Fastsurance,
-- Vitra.ai, FinBox, and Castler were wrongly defaulted into fund-2
-- earlier in this project and actually belong in fund-1 (fund-2 rows
-- for them are removed here, not just supplemented).
-- ============================================================

-- Remove the incorrect fund-2 memberships for companies that are
-- actually fund-1-only (or fund-1-and-fund-2, handled below).
delete from fund_companies
where fund_id = (select id from funds where slug = 'fund-2')
  and company_id in (
    select id from companies where slug in ('fastsurance', 'apexFutureLabs', 'finbox', 'castler')
  );

-- Fund 1: full correct membership.
insert into fund_companies (fund_id, company_id, display_order)
select f.id, c.id, row_number() over (order by c.slug)
from funds f, companies c
where f.slug = 'fund-1'
  and c.slug in (
    'easyrewardz', 'grayquest', 'riskcovry', 'multipl',
    'fastsurance', 'apexFutureLabs', 'leegality', 'finbox', 'castler'
  )
on conflict (fund_id, company_id) do nothing;

-- Fund 2: add Leegality and GrayQuest (they span both funds).
-- Fundamento/Knight FinTech/Traqcheck were already correctly fund-2-only.
insert into fund_companies (fund_id, company_id, display_order)
select f.id, c.id, row_number() over (order by c.slug)
from funds f, companies c
where f.slug = 'fund-2'
  and c.slug in ('leegality', 'grayquest')
on conflict (fund_id, company_id) do nothing;
