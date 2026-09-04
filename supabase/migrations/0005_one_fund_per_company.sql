-- ============================================================
-- One fund per company.
--
-- "Fund details.xlsx" (the authoritative allocation, also used by
-- 0003_fund_corrections.sql) lists three companies under BOTH funds —
-- Leegality, GrayQuest and Finvu — because Fund 2 followed on into
-- companies Fund 1 had already backed. 0003 reproduced that faithfully,
-- which is why the dashboard's company picker showed Leegality and
-- GrayQuest under Fund 1 AND Fund 2.
--
-- User decision: a company appears under exactly ONE fund. Each
-- duplicate is assigned to the fund it originated in, which is Fund 1
-- for all three:
--   * the Excel lists them in the Fund 1 block first, and
--   * both carry a full per-company slide in the Fund 1 newsletter
--     ("Fund 1 Newsletter - Q4FY26.pptx", the deck the generated
--     reports are cloned from).
--
-- Resulting membership (Finvu has no `companies` row yet, so it is not
-- represented either way):
--   Fund 1: Leegality, FinBox, Data Sutram, EasyRewardz, Multipl,
--           Insurance Samadhan, Riskcovry, Castler, GrayQuest, Vitra.ai
--   Fund 2: Fundamento, Knight FinTech, Traqcheck
--
-- This drops fund_companies rows only — no company, config, upload or
-- metric row is touched, so nothing about a company's data or dashboard
-- changes, only which fund lists it. Reversible by re-running 0003.
-- Safe to re-run.
-- ============================================================

delete from fund_companies
where fund_id = (select id from funds where slug = 'fund-2')
  and company_id in (
    select id from companies where slug in ('leegality', 'grayquest')
  );

-- Close the gaps left in fund-2's display_order by the deletions, so the
-- picker's ordering stays 1..n rather than skipping numbers.
with renumbered as (
  select fc.fund_id,
         fc.company_id,
         row_number() over (partition by fc.fund_id order by fc.display_order, c.slug) as new_order
  from fund_companies fc
  join companies c on c.id = fc.company_id
)
update fund_companies fc
set display_order = r.new_order
from renumbered r
where fc.fund_id = r.fund_id
  and fc.company_id = r.company_id
  and fc.display_order is distinct from r.new_order;
