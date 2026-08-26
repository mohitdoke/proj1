-- ============================================================
-- Reference data seed: funds, companies, fund membership.
-- Pure static reference data (names/slugs) — NOT financial MIS data,
-- which is imported separately via `npm run seed:mis` (scripts/seed.mjs)
-- from the actual master Excel files, never typed by hand.
--
-- Fund membership below is a STARTING POINT (Fund 1 = the four
-- companies onboarded first, Fund 2 = the five added after) — update
-- fund_companies once the real IIFL Fintech Fund 1 / Fund 2 membership
-- is confirmed. Reassigning a company to a different fund is a
-- one-row UPDATE here, no code change.
--
-- Safe to re-run: every insert is an idempotent upsert.
-- ============================================================

insert into funds (name, slug) values
  ('IIFL Fintech Fund 1', 'fund-1'),
  ('IIFL Fintech Fund 2', 'fund-2')
on conflict (slug) do nothing;

insert into companies (slug, name, legal_name) values
  ('easyrewardz',    'Easyrewardz',        'Easyrewardz Software Services Private Limited'),
  ('grayquest',      'GrayQuest',          'GrayQuest Education Finance Private Limited'),
  ('riskcovry',      'Riskcovry',          'Riskcovry'),
  ('multipl',        'Multipl',            'Multipl Fintech Solutions Private Limited'),
  ('fastsurance',    'Insurance Samadhan', 'FASTSURANCE Consultants Private Limited'),
  ('apexFutureLabs', 'Vitra.ai',           'Apex Future Labs Private Limited'),
  ('leegality',      'Leegality',          'Grey Swift Private Limited'),
  ('finbox',         'FinBox',             'MOSHPIT Technologies Private Limited'),
  ('fundamento',     'Fundamento',         'Fundamento')
on conflict (slug) do update set
  name = excluded.name,
  legal_name = excluded.legal_name;

-- One company_configs row per company, pointing at its code-side
-- COMPANY_CONFIGS key (identical to the slug for every company today).
insert into company_configs (company_id, config_key)
select id, slug from companies
on conflict (company_id) do update set config_key = excluded.config_key;

-- Fund 1: the four companies onboarded first.
insert into fund_companies (fund_id, company_id, display_order)
select f.id, c.id, row_number() over (order by c.slug)
from funds f, companies c
where f.slug = 'fund-1'
  and c.slug in ('easyrewardz', 'grayquest', 'riskcovry', 'multipl')
on conflict (fund_id, company_id) do nothing;

-- Fund 2: the five companies added afterwards.
insert into fund_companies (fund_id, company_id, display_order)
select f.id, c.id, row_number() over (order by c.slug)
from funds f, companies c
where f.slug = 'fund-2'
  and c.slug in ('fastsurance', 'apexFutureLabs', 'leegality', 'finbox', 'fundamento')
on conflict (fund_id, company_id) do nothing;
