with maintenance_task as (
  select
    coalesce(
      (
        select task ->> 'id'
        from public.app_settings settings
        cross join lateral jsonb_array_elements(coalesce(settings.payload -> 'tasks', '[]'::jsonb)) as task
        where settings.setting_key = 'maintenance'
          and lower(coalesce(task ->> 'task', '')) = 'bedbugs'
        limit 1
      ),
      'maintenance-task-10-bedbugs'
    ) as task_id,
    'Bedbugs'::text as task_name
),
source(where_value, done_date, type, who, note) as (
  values
    ('102', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('105', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('201', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('202', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('203', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('204', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('205', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('206', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('207', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('211', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('212', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('213', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('214', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('215', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('216', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('217', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('218', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('111', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('112', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('113', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('114', '2024-01-13', 'Periodica', 'CentrosPragas', ''),
    ('105', '2023-12-27', 'Desinfestação', 'CentrosPragas', ''),
    ('102', '2024-01-19', 'Desinfestação', 'CentrosPragas', 'Nao se verificou presença'),
    ('105', '2024-01-19', 'Desinfestação', 'CentrosPragas', 'Nao se verificou presença'),
    ('102', '2024-07-09', 'Desinfestação', 'CentrosPragas', 'Desinfestação e desmontagem das camas pelos Filipes. Foram encontrados bastantes'),
    ('102', '2024-07-16', 'Desinfestação', 'CentrosPragas', '2ª desinfestação, ainda foram encontrados'),
    ('102', '2024-08-27', 'Desinfestação', 'CentrosPragas', ''),
    ('102', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('105', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('201', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('202', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('203', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('204', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('205', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('206', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('207', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('211', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('212', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('213', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('214', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('215', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('216', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('217', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('218', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('111', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('112', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('113', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('114', '2025-01-14', 'Periodica', 'CentrosPragas', ''),
    ('3E', '2025-03-19', 'Desinfestação', 'CentrosPragas', 'foram encontrados alguns no primeiro quarto'),
    ('3E', '2023-03-31', 'Desinfestação', 'CentrosPragas', '2ª desinfestação, nao foram detetados'),
    ('216', '2025-06-23', 'Desinfestação', 'CentrosPragas', 'Havia bastantes'),
    ('215', '2025-06-23', 'Desinfestação', 'CentrosPragas', 'Por precaução e estar o quarto disponivel. Não foram detetados'),
    ('216', '2025-07-16', 'Desinfestação', 'CentrosPragas', '2ª desinfestação. Já não foram detectados'),
    ('206', '2025-07-16', 'Desinfestação', 'CentrosPragas', 'Desinfestação foram detectados bastantes'),
    ('206', '2025-07-28', 'Desinfestação', 'CentrosPragas', '2ª desinfestação. Foram desmontada as camas pelos Filipes. Ainda foram detetados bastantes junto a cama 1 e 2'),
    ('206', '2025-10-02', 'Desinfestação', 'CentrosPragas', 'Detectados junto a cama 3,4 e 5,6'),
    ('206', '2025-10-20', 'Desinfestação', 'CentrosPragas', '2ª desinfestação - ainda foi detectada presença junto a cama 6'),
    ('203', '2025-11-12', 'Desinfestação', 'CentrosPragas', 'Não foram detectados'),
    ('203', '2025-11-28', 'Desinfestação', 'CentrosPragas', '2ª desinfestação. Não foram detectados'),
    ('102', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('105', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('201', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('202', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('203', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('204', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('205', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('206', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('207', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('211', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('212', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('213', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('214', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('215', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('216', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('217', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('218', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('111', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('112', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('113', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('114', '2026-01-12', 'Periodica', 'CentrosPragas', ''),
    ('206', '2026-04-16', 'Desinfestação', 'CentrosPragas', 'Encontrados bastantes nas camas 2 e 4'),
    ('206', '2026-04-16', 'Intervenção', 'Miguel', 'Deitadas fora e substituidas camas 2 e 4. Lavadas Cortinas e todas as roupas de cama'),
    ('206', '2026-04-30', 'Desinfestação', 'CentrosPragas', 'Não foram detectados'),
    ('217', '2023-07-08', 'Intervenção', 'Miguel', 'encontrado um vivo na cama 3'),
    ('217', '2023-07-09', 'Intervenção', 'Miguel', 'Inspeção e colocação de bomba inseticida - nao foram detectados mais'),
    ('217', '2023-07-31', 'Desinfestação', 'CentrosPragas', ''),
    ('218', '2023-08-05', 'Suspeita', 'Miguel', 'Verificamos o quarto todo depois de suspeita de picadas, mas não foi nada encontrado')
)
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
  task.task_id,
  task.task_name,
  source.where_value,
  source.done_date::date,
  source.type,
  source.who,
  source.note
from source
cross join maintenance_task as task
where not exists (
  select 1
  from public.maintenance_logs existing
  where existing.task_id = task.task_id
    and existing.where_value = source.where_value
    and existing.done_date = source.done_date::date
    and existing.type = source.type
    and existing.who = source.who
    and existing.note = source.note
);
