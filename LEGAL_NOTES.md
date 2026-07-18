# Minor scheduling checks

Last reviewed: July 17, 2026.

The app applies the stricter applicable Tennessee or federal scheduling limit. Primary references:

- Tennessee Department of Labor & Workforce Development, Child Labor Act: https://www.tn.gov/workforce/employees/labor-laws/labor-laws-redirect/child-labor.html
- U.S. Department of Labor, Fact Sheet #43: https://www.dol.gov/agencies/whd/fact-sheets/43-child-labor-non-agriculture

## Checks implemented

For 14- and 15-year-olds:

- Work must be outside the employee's entered school hours.
- School day: maximum 3 hours.
- School week: maximum 18 hours.
- Nonschool day: maximum 8 hours.
- Nonschool week: maximum 40 hours.
- Work may not begin before 7:00 AM.
- Work may not end after 7:00 PM, except that the evening limit is 9:00 PM from June 1 through Labor Day when school is not in session.

For 16- and 17-year-olds:

- Work must be outside the employee's entered school hours.
- Sunday through Thursday before a school day: work may not continue after 10:00 PM without a valid parental consent form.
- With the consent form recorded, work may continue until midnight on no more than three such nights per week.

For every employee under 18:

- A shift of six consecutive hours produces a reminder that Tennessee requires a 30-minute unpaid break, not during or before the first hour.
- An entry without a definite end time produces a warning that compliance cannot be verified.
- The app aggregates all roles and all split-shift segments for the same employee and flags overlapping assignments.

## Not implemented

- The app does not determine whether a minor's assigned duties or equipment are legally permitted. Restaurants must separately enforce occupation and equipment restrictions.
- It does not know a school's calendar or class times automatically. A manager must maintain the week-level school-day selections and employee-specific school hours.
- It does not validate work-based-learning, homeschool, church-related-school, or other statutory exceptions.
- It does not replace required personnel records or legal review.
- Historical POS clock-out data is not treated as the scheduled end time. It may later be used to suggest common shift ranges, but the manager must choose the actual scheduled end time.
