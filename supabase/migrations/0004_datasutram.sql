-- ============================================================
-- Add Data Sutram (Extrapolate Advisors Private Limited) as a new
-- company, per Fund details.xlsx: Fund 1 only.
-- Safe to re-run: every insert is an idempotent upsert.
-- ============================================================

insert into companies (slug, name, legal_name) values
  ('datasutram', 'Data Sutram', 'Extrapolate Advisors Private Limited')
on conflict (slug) do update set
  name = excluded.name,
  legal_name = excluded.legal_name;

insert into company_configs (company_id, config_key)
select id, slug from companies where slug = 'datasutram'
on conflict (company_id) do update set config_key = excluded.config_key;

insert into fund_companies (fund_id, company_id, display_order)
select f.id, c.id, (select coalesce(max(display_order), 0) + 1 from fund_companies where fund_id = f.id)
from funds f, companies c
where f.slug = 'fund-1'
  and c.slug = 'datasutram'
on conflict (fund_id, company_id) do nothing;
