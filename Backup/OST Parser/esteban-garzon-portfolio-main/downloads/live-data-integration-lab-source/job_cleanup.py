FIND_DUPLICATES_QUERY = """
SELECT LOWER(TRIM(job_name)) AS normalized_name, COUNT(*)
FROM job_data
GROUP BY LOWER(TRIM(job_name))
HAVING COUNT(*) > 1;
"""


DELETE_DUPLICATE_QUERY = """
DELETE FROM job_data
WHERE ctid IN (
    SELECT ctid
    FROM job_data
    WHERE LOWER(TRIM(job_name)) = LOWER(TRIM(%s))
      AND ctid NOT IN (
          SELECT MIN(ctid)
          FROM job_data
          WHERE LOWER(TRIM(job_name)) = LOWER(TRIM(%s))
      )
);
"""
