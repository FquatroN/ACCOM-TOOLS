-- Source: Feature 12 - Migracao maitenance_AQS.xlsx
-- Loads AQS maintenance history into public.maintenance_logs.
-- Resolves task_id from the saved Maintenance settings by task name.
-- Skips rows that already exist in maintenance_logs.

create temp table maintenance_import_source (
  task_name text not null,
  where_value text not null,
  done_date date not null,
  type text not null,
  who text not null,
  note text not null default ''
) on commit drop;

insert into maintenance_import_source (task_name, where_value, done_date, type, who, note)
values
  ('AQS', '10 Termoacumulador', '2023-04-12'::date, 'Novo', 'Catagas', 'Novo montado por Catagas. Anterior rompeu'),
  ('AQS', '10 Termoacumulador', '2025-04-11'::date, 'Manutenção', 'Catagas', 'Manutenção Periodica. Anodo foi substituido'),
  ('AQS', '10 Termoacumulador', '2026-04-10'::date, 'Avaria', 'Tiago', 'Detectada avaria. Está a pingar água e a mandar o quadro a baixo por vezes'),
  ('AQS', '10 Termoacumulador', '2026-04-26'::date, 'Novo', 'Miguel', 'Novo montado por Miguel'),
  ('AQS', '10 CaldeiraDeposito', '2024-01-09'::date, 'Reparação', 'Catagas', 'Estava a perder pressão e levou nova valvula de segurança'),
  ('AQS', '10 CaldeiraDeposito', '2024-06-12'::date, 'Reparação', 'Catagas', 'Nova bomba e painel de controlo'),
  ('AQS', '10 CaldeiraDeposito', '2025-03-21'::date, 'Novo', 'Miguel', 'Levou deposito comprado por Miguel, pois Catagas não sabia quando teria disponivel e qd seria feita montagem. Isto depois de esperar 2 semanas por visita para ver o que era o problema'),
  ('AQS', '10 CaldeiraDeposito', '2025-06-23'::date, 'Reparação', 'Catagas', 'Estava a perder pressão e fazer barulho quando arrancava. Tinha estabilizado, mas Catagas verificou e fez qq coisa na valvula de 3 vias'),
  ('AQS', '10 CaldeiraDeposito', '2026-03-26'::date, 'Manutenção', 'EnergiaViva', 'Manutenção, Esta a perder pressoa. Energia viva verificou e disse que era o Vaso de expansão. A aguardar que venham trocar'),
  ('AQS', '10 CaldeiraDeposito', '2026-04-01'::date, 'Reparação', 'EnergiaViva', 'Foi trocado o vaso de Expansão pela Energia Viva.'),
  ('AQS', '11 StaffTermoacumulador', '2025-02-26'::date, 'Novo', 'Catagas', 'Novo Termoacumulador. Anterior tinha rompido'),
  ('AQS', '112 Termoacumulador', '2022-11-11'::date, 'Novo', 'Leonel', 'Novo Termoacumulador. Anterior tinha rompido'),
  ('AQS', '20 CaldeiraDeposito', '2023-06-20'::date, 'Reparação', 'Catagas', 'Nova bomba e limitdor de temperatura'),
  ('AQS', '20 CaldeiraDeposito', '2024-09-04'::date, 'Reparação', 'Catagas', 'estava a perder pressão rapidamente, levou peça nova'),
  ('AQS', '20 CaldeiraDeposito', '2025-02-26'::date, 'Manutenção', 'Catagas', 'Manutenção periodica, mas sem grande detalhe'),
  ('AQS', '20 CaldeiraDeposito', '2026-02-19'::date, 'Reparação', 'Fernando', 'Valvula termostatica avariou. Fernando veio substituir'),
  ('AQS', '20 CaldeiraDeposito', '2026-04-30'::date, 'Manutenção', 'EnergiaViva', 'Mantenção periodica'),
  ('AQS', '21 CaldeiraDeposito', '2023-08-07'::date, 'Reparação', 'Catagas', 'Levou bomba de circulação nova'),
  ('AQS', '21 CaldeiraDeposito', '2024-09-09'::date, 'Reparação', 'Catagas', 'Estava a perder pressão. Substituida valvula'),
  ('AQS', '21 CaldeiraDeposito', '2025-02-11'::date, 'Reparação', 'Catagas', 'Estava a dar erro F0, nova placa eletronica'),
  ('AQS', '21 CaldeiraDeposito', '2025-04-16'::date, 'Novo', 'Catagas', 'Novo depósito. O anterior tinha rompido'),
  ('AQS', '21 CaldeiraDeposito', '2025-06-23'::date, 'Reparação', 'Catagas', 'Estava a perder pressão e fazer barulho quando arrancava. Tinha estabilizado, mas Catagas verificou e fez qq coisa na valvula de 3 vias'),
  ('AQS', '21 CaldeiraDeposito', '2026-04-30'::date, 'Manutenção', 'EnergiaViva', 'Mantenção periodica'),
  ('AQS', '21 CaldeiraDeposito', '2026-05-02'::date, 'Reparação', 'EnergiaViva', 'Caldeira parou. detectamos que era o fio do sensor solto. Tiago arranjou, mas depois botão de ligar não funcionava. Energia viva veio de manhã dia 02 para reparar'),
  ('AQS', '2D CaldeiraDeposito', '2026-04-30'::date, 'Manutenção', 'EnergiaViva', 'Mantenção periodica'),
  ('AQS', '2E CaldeiraDeposito', '2023-09-13'::date, 'Reparação', 'Catagas', 'Caldeira levou nova bomba'),
  ('AQS', '2E CaldeiraDeposito', '2024-04-24'::date, 'Reparação', 'Catagas', 'Estava a dar erro E9, Aperto de tampa da placa'),
  ('AQS', '2E CaldeiraDeposito', '2025-02-11'::date, 'Reparação', 'Catagas', 'Estava a dar erro E9, Vai precisar de nova boma'),
  ('AQS', '2E CaldeiraDeposito', '2025-02-26'::date, 'Manutenção', 'Catagas', 'Levou nova bomba, foi feita manutenção à Caldeira apenas'),
  ('AQS', '2E CaldeiraDeposito', '2025-04-04'::date, 'Reparação', 'EnergiaViva', 'Caldeira Avariou com erro E9. Catagas veio e nao tinha bomba para substituir. Chamado Hugo Costa que substitui bomba - levou 360€'),
  ('AQS', '2E CaldeiraDeposito', '2025-05-25'::date, 'Reparação', 'EnergiaViva', 'Caldeira estava desligada sem ligar. Tivemos que chamar outra empresa. Foi só substituir um fusivel'),
  ('AQS', '2E CaldeiraDeposito', '2026-03-05'::date, 'Reparação', 'EnergiaViva', 'Estava a perder pressão. Foi retificada a valvula'),
  ('AQS', '3E CaldeiraDeposito', '2026-03-28'::date, 'Manutenção', 'EnergiaViva', 'Manutenção pela EnergiaViva - substituido anodo'),
  ('AQS', '3E CaldeiraDeposito', '2026-04-30'::date, 'Alteração', 'EnergiaViva', 'Montada resistência eletrica no deposito'),
  ('AQS', '4D CaldeiraDeposito', '2022-08-11'::date, 'Reparação', 'Catagas', 'Bomba de circulação e valvula de segurança'),
  ('AQS', '4D CaldeiraDeposito', '2024-09-09'::date, 'Reparação', 'Catagas', 'Estava a perder pressão, Substituida valvula'),
  ('AQS', '4D CaldeiraDeposito', '2025-02-26'::date, 'Manutenção', 'Catagas', 'Manutenção light'),
  ('AQS', '4D CaldeiraDeposito', '2025-04-11'::date, 'Reparação', 'Catagas', 'Estava a perder pressao, levou novo vaso de expansao'),
  ('AQS', '4D CaldeiraDeposito', '2025-06-02'::date, 'Reparação', 'Catagas', 'Lubrificação da bomba de circulação - caldeira estava a ficar sem pressão'),
  ('AQS', '4D CaldeiraDeposito', '2025-06-20'::date, 'Reparação', 'Catagas', 'Estava a perder pressão, mas depois estabilizou, fou apenas verificada pela Catagas'),
  ('AQS', '4D CaldeiraDeposito', '2025-09-15'::date, 'Novo', 'Fernando', 'Deposito rompeu, Miguel comprou novo deposito e Fernando montou'),
  ('AQS', '4D CaldeiraDeposito', '2026-03-24'::date, 'Manutenção', 'EnergiaViva', 'Manutenção pela Energia Viva'),
  ('AQS', '4E CaldeiraDeposito', '2025-02-26'::date, 'Manutenção', 'Catagas', 'Manutenção periodica'),
  ('AQS', '5E Esquentador', '2022-10-20'::date, 'Novo', 'Leonel', 'Esquentador nao funcionava. Problema era na chaminé entupida, mas entretanto comprámos um esquentador novo');

do $$
declare
  missing_tasks text;
begin
  select string_agg(task_name, ', ' order by task_name)
    into missing_tasks
  from (
    select distinct source.task_name
    from maintenance_import_source as source
    left join (
      select
        task ->> 'id' as task_id,
        task ->> 'task' as task_name
      from public.app_settings settings
      cross join lateral jsonb_array_elements(coalesce(settings.payload -> 'tasks', '[]'::jsonb)) as task
      where settings.setting_key = 'maintenance'
    ) as task_lookup
      on lower(trim(task_lookup.task_name)) = lower(trim(source.task_name))
    where task_lookup.task_id is null
  ) as missing;

  if missing_tasks is not null then
    raise exception 'Maintenance task names missing from saved settings: %', missing_tasks;
  end if;
end $$;

insert into public.maintenance_logs (
  task_id,
  task_name,
  where_value,
  done_date,
  type,
  who,
  note
)
select
  task_lookup.task_id,
  task_lookup.task_name,
  source.where_value,
  source.done_date,
  source.type,
  source.who,
  source.note
from maintenance_import_source as source
join (
  select
    task ->> 'id' as task_id,
    task ->> 'task' as task_name
  from public.app_settings settings
  cross join lateral jsonb_array_elements(coalesce(settings.payload -> 'tasks', '[]'::jsonb)) as task
  where settings.setting_key = 'maintenance'
) as task_lookup
  on lower(trim(task_lookup.task_name)) = lower(trim(source.task_name))
where not exists (
  select 1
  from public.maintenance_logs existing
  where existing.task_id = task_lookup.task_id
    and existing.where_value = source.where_value
    and existing.done_date = source.done_date
    and existing.type = source.type
    and existing.who = source.who
    and existing.note = source.note
);
