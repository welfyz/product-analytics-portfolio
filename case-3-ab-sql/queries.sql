-- =====================================================================
-- Кейс: A/B-тест новой формы оплаты ЖКХ (SQLite)
-- Таблицы:
--   users        (user_id, ab_group, age, city, device_type)   -- 1 500 строк, очищенные данные
--   payments     (payment_id, user_id, ab_group, step1_opened,
--                 step2_entered, step3_confirmed, step4_success) -- 2 657 строк, очищенные данные
--   users_raw, payments_raw -- те же таблицы до очистки (2 660 платежей)
-- Конверсия считается по людям: успех = хотя бы одна успешная оплата.
-- Ожидаемые результаты указаны в комментариях и совпадают с расчётом в Excel.
-- =====================================================================


-- 0. Поиск аномальных записей в сырых данных
-- Ожидается 3 строки: 7658 (нет пользователя), 99001 (интервал 1 минута),
-- 99002 (группа не совпала)
SELECT p.payment_id, p.user_id,
       p.ab_group AS payment_group,
       u.ab_group AS user_group,
       CASE
         WHEN u.user_id IS NULL THEN 'пользователя нет в списке'
         WHEN p.ab_group <> u.ab_group THEN 'группа не совпала с группой пользователя'
         ELSE 'аномальный интервал шаг 3 → шаг 4'
       END AS reason
FROM payments_raw p
LEFT JOIN users_raw u ON p.user_id = u.user_id
WHERE u.user_id IS NULL
   OR p.ab_group <> u.ab_group
   OR (p.step4_success IS NOT NULL
       AND strftime('%s', p.step4_success) - strftime('%s', p.step3_confirmed) < 120)
ORDER BY p.payment_id;


-- 1. Проверка качества очищенных данных (везде должно быть 0)
SELECT 'платежи без пользователя' AS check_name, COUNT(*) AS bad_rows
FROM payments p
LEFT JOIN users u ON p.user_id = u.user_id
WHERE u.user_id IS NULL

UNION ALL
SELECT 'группа платежа не совпала с группой пользователя', COUNT(*)
FROM payments p
JOIN users u ON p.user_id = u.user_id
WHERE p.ab_group <> u.ab_group

UNION ALL
SELECT 'дубли payment_id', COUNT(*) - COUNT(DISTINCT payment_id)
FROM payments

UNION ALL
SELECT 'интервал шаг 3 → шаг 4 меньше 2 минут', COUNT(*)
FROM payments
WHERE step4_success IS NOT NULL
  AND strftime('%s', step4_success) - strftime('%s', step3_confirmed) < 120;


-- 2. Воронка по шагам: уникальные пользователи, дошедшие до шага
-- Ожидается: A 750 / 670 / 546 / 491, B 750 / 707 / 644 / 603
SELECT ab_group,
       COUNT(DISTINCT CASE WHEN step1_opened    IS NOT NULL THEN user_id END) AS step1,
       COUNT(DISTINCT CASE WHEN step2_entered   IS NOT NULL THEN user_id END) AS step2,
       COUNT(DISTINCT CASE WHEN step3_confirmed IS NOT NULL THEN user_id END) AS step3,
       COUNT(DISTINCT CASE WHEN step4_success   IS NOT NULL THEN user_id END) AS step4
FROM payments
GROUP BY ab_group;


-- 3. Общая конверсия шаг 1 → шаг 4 по группам
-- Ожидается: A 65.5%, B 80.4%
WITH f AS (
  SELECT ab_group,
         COUNT(DISTINCT user_id) AS opened,
         COUNT(DISTINCT CASE WHEN step4_success IS NOT NULL THEN user_id END) AS paid
  FROM payments
  GROUP BY ab_group
)
SELECT ab_group, opened, paid,
       ROUND(100.0 * paid / opened, 1) AS conversion_pct
FROM f;


-- 4. Разница между группами: в процентных пунктах и в процентах
-- Ожидается: 14.9 п.п., +22.8%
WITH f AS (
  SELECT ab_group,
         COUNT(DISTINCT user_id) AS opened,
         COUNT(DISTINCT CASE WHEN step4_success IS NOT NULL THEN user_id END) AS paid
  FROM payments
  GROUP BY ab_group
),
c AS (
  SELECT MAX(CASE WHEN ab_group = 'A' THEN 1.0 * paid / opened END) AS a,
         MAX(CASE WHEN ab_group = 'B' THEN 1.0 * paid / opened END) AS b
  FROM f
)
SELECT ROUND(100 * (b - a), 1)     AS diff_pp,
       ROUND(100 * (b - a) / a, 1) AS rel_pct
FROM c;


-- 5. Конверсия по группам и устройствам (проверка парадокса Симпсона)
-- Ожидается: android A 64.3 / B 79.9, ios A 66.9 / B 80.8
SELECT p.ab_group, u.device_type,
       COUNT(DISTINCT p.user_id) AS opened,
       COUNT(DISTINCT CASE WHEN p.step4_success IS NOT NULL THEN p.user_id END) AS paid,
       ROUND(100.0 * COUNT(DISTINCT CASE WHEN p.step4_success IS NOT NULL THEN p.user_id END)
             / COUNT(DISTINCT p.user_id), 1) AS conversion_pct
FROM payments p
JOIN users u ON p.user_id = u.user_id
GROUP BY p.ab_group, u.device_type
ORDER BY u.device_type, p.ab_group;


-- 6. Конверсия между соседними шагами воронки
-- Ожидается: A 89.3 / 81.5 / 89.9, B 94.3 / 91.1 / 93.6
WITH f AS (
  SELECT ab_group,
         COUNT(DISTINCT CASE WHEN step1_opened    IS NOT NULL THEN user_id END) AS step1,
         COUNT(DISTINCT CASE WHEN step2_entered   IS NOT NULL THEN user_id END) AS step2,
         COUNT(DISTINCT CASE WHEN step3_confirmed IS NOT NULL THEN user_id END) AS step3,
         COUNT(DISTINCT CASE WHEN step4_success   IS NOT NULL THEN user_id END) AS step4
  FROM payments
  GROUP BY ab_group
)
SELECT ab_group,
       ROUND(100.0 * step2 / step1, 1) AS s1_to_s2,
       ROUND(100.0 * step3 / step2, 1) AS s2_to_s3,
       ROUND(100.0 * step4 / step3, 1) AS s3_to_s4
FROM f;


-- 7. Z-тест для двух долей (считаем z в квадрате, корень в SQLite не нужен)
-- Ожидается: z_squared около 42.36 (z = 6.51). Больше 10.83 значит p < 0.001
WITH f AS (
  SELECT ab_group,
         COUNT(DISTINCT user_id) AS n,
         COUNT(DISTINCT CASE WHEN step4_success IS NOT NULL THEN user_id END) AS x
  FROM payments
  GROUP BY ab_group
),
c AS (
  SELECT MAX(CASE WHEN ab_group = 'A' THEN n END) AS n_a,
         MAX(CASE WHEN ab_group = 'A' THEN x END) AS x_a,
         MAX(CASE WHEN ab_group = 'B' THEN n END) AS n_b,
         MAX(CASE WHEN ab_group = 'B' THEN x END) AS x_b
  FROM f
),
p AS (
  SELECT n_a, n_b,
         1.0 * x_a / n_a AS p_a,
         1.0 * x_b / n_b AS p_b,
         1.0 * (x_a + x_b) / (n_a + n_b) AS p_pool
  FROM c
)
SELECT ROUND(p_a, 4) AS p_a,
       ROUND(p_b, 4) AS p_b,
       ROUND((p_b - p_a) * (p_b - p_a)
             / (p_pool * (1 - p_pool) * (1.0 / n_a + 1.0 / n_b)), 2) AS z_squared
FROM p;
