from datetime import date, timedelta

import otl
from staff import GroupMember

# Relevant codes for the group
magnet_lab = otl.Code('STGA00029', '02', priority=otl.Priority.BALANCING)
mars_underpinning = otl.Code('STGA00206', priority=otl.Priority.BALANCING)
clara_scientific = otl.Code('STGA09000', '520',
                            end=date(2025, 11, 30), priority=otl.Priority.AGREED)
clara_user_facility = otl.Code('STGA00265',
                               start=clara_scientific.end + timedelta(days=1), priority=otl.Priority.AGREED)
# was Apr-Aug, 5mths; Sep-Mar, 7mths
# now Apr-Nov, 8mths; Dec-Mar, 4mths
scu = otl.Code('STGA00273', priority=otl.Priority.AGREED)
sustainable_accelerators = otl.Code('STGA00298', priority=otl.Priority.BALANCING)
xfel_rnd = otl.Code('STGA00241', priority=otl.Priority.BALANCING)
novel_acceleration = otl.Code('STGA00242', priority=otl.Priority.BALANCING)
ai_ml = xfel_rnd  # Code('STGA09999')  # no code yet
thin_films = otl.Code('STGA00501', priority=otl.Priority.BALANCING)
novel_neg = otl.Code('STGA00502', priority=otl.Priority.BALANCING)
dae = otl.Code('STGA02000', '03', end=date(2026, 3, 31))  # now runs to end FY25/26
ruedi_epsrc_bridging = otl.Code('STGA02006', '02', end=date(2025, 6, 30))
ruedi_post_bridging = otl.Code('STGA02008', start=date(2025, 7, 1),
                               end=date(2025, 10, 31), priority=otl.Priority.AGREED)
# code changed in November (email from Julian McKenzie 29/10/25)
ruedi_new_code = otl.Code('STGA02011', '02', start=date(2025, 11, 1), priority=otl.Priority.AGREED)
zepto_clara_gott = otl.Code('STGA02005', end=date(2026, 2, 28))
clepto_pocf = otl.Code('STLA00037', '147', start=date(2025, 12, 1))

def ukxfel_cdoa(wp: int):
    return otl.Code('STKA00183', f'03.{wp:02d}', end=date(2025, 9, 30))


ukxfel_continuation = otl.Code('STGA00183', '01', start=date(2025, 10, 1))

epac = otl.Code('STKA01103', '06.01')

members: list[GroupMember] = [
    GroupMember('Ben Shepherd', 207835,
                person_id=100000020410836, assignment_id=300000117877863,
                booking_plan=otl.BookingPlan([
                    otl.Entry(epac, 0.05),
                    otl.Entry(dae, 0.08),
                    otl.Entry(ruedi_new_code, 0.0434),
                    otl.Entry(clara_user_facility, 0.25),
                    otl.Entry(sustainable_accelerators, 0.46),
                    otl.Entry(magnet_lab),
                ])),
    GroupMember('Alexander Bainbridge',
                200394,
                'alex.bainbridge@stfc.ac.uk', known_as='Alex B',
                person_id=100000020410917, assignment_id=300000117882174,
                booking_plan=otl.BookingPlan([
                    otl.Entry(epac, 0.1),
                    otl.Entry(ruedi_epsrc_bridging, 0.05),
                    otl.Entry(ruedi_new_code, 0.1851),
                    otl.Entry(clara_scientific, 0.05),
                    otl.Entry(clara_user_facility, 0.2),
                    otl.Entry(magnet_lab),
                ])),
    GroupMember('David Dunning',
                204991,
                'david.dunning@stfc.ac.uk', known_as='Dave',
                person_id=100000020417326, assignment_id=300000117978650,
                booking_plan=otl.BookingPlan([
                    otl.Entry(ukxfel_cdoa(wp=1), 0.175),
                    otl.Entry(ukxfel_cdoa(wp=5), 0.175),
                    otl.Entry(ukxfel_continuation, 0.4),
                    otl.Entry(xfel_rnd, 0.15),
                    otl.Entry(ai_ml),
                ])),
    GroupMember('Neil Thompson',
                206988,
                person_id=100000020415442, assignment_id=300000117997606,
                booking_plan=otl.BookingPlan([
                    otl.Entry(epac, 0.2),
                    otl.Entry(ukxfel_cdoa(wp=5), 0.2),
                    otl.Entry(ukxfel_continuation, 0.25),
                    otl.Entry(xfel_rnd, 0.25),
                    otl.Entry(novel_acceleration),
                ])),
    GroupMember('Kiril Marinov',
                204936,
                person_id=100000020410826, assignment_id=300000117884336,
                booking_plan=otl.BookingPlan([
                    otl.Entry(novel_neg, 0.25),
                    otl.Entry(thin_films, 0.5),
                    otl.Entry(mars_underpinning),
                ])),
    GroupMember('Alan Mak',
                206367,
                person_id=100000020412139, assignment_id=300000117893427,
                booking_plan=otl.BookingPlan([
                    otl.Entry(ukxfel_cdoa(wp=5), 0.5),
                    otl.Entry(ukxfel_continuation, 0.25),
                    otl.Entry(xfel_rnd),
                ])),
    GroupMember('Alexander Hinton',
                201375,
                'alex.hinton@stfc.ac.uk', known_as='Alex H',
                person_id=100000020413904, assignment_id=300000117923738,
                booking_plan=otl.BookingPlan([
                    otl.Entry(dae, 0.3),
                    otl.Entry(zepto_clara_gott, 0.08),
                    otl.Entry(clepto_pocf, 168 / otl.hours_per_fte),
                    otl.Entry(scu),
                ]),
                ignore_days={
                    date(2025, 6, 19),
                    date(2025, 6, 20),
                    date(2025, 6, 23),
                    date(2025, 6, 24),
                }),
    GroupMember('Amelia Pollard',
                205179,
                person_id=100000020414057, assignment_id=300000117928903,
                known_as='Amy',
                booking_plan=otl.BookingPlan([
                    otl.Entry(ai_ml),
                ])),
    GroupMember('Matthew King',
                207007,
                'matthew.king@stfc.ac.uk', known_as='Matt',
                person_id=100000020413933, assignment_id=300000117929802,
                booking_plan=otl.BookingPlan([
                    otl.Entry(clara_scientific, 0.17),
                    otl.Entry(clara_user_facility, 0.35),
                    otl.Entry(ai_ml),
                ])),
    GroupMember('Nasiq Ziyan',
                207521,
                person_id=100000020417760, assignment_id=300000117981407,
                booking_plan=otl.BookingPlan([
                    otl.Entry(clara_scientific, 0.17),
                    otl.Entry(clara_user_facility, 0.35),
                    otl.Entry(ai_ml),
                ])),
    # GroupMember('Tom Smith', 0, 100000020412335, 300000117987914, 'thomas.smith@stfc.ac.uk'),
]


# if __name__ == '__main__':
# check_total_ftes(members)
# print(*[person.name for person in members], sep='\t')
# person.update_off_days()
# print(*sorted(list(person.off_days)), sep='\n')
# print(person.daily_bookings(date.today()))
# run_otl_calculator()
#     for entry in member.booking_plan.entries:
#         print(entry.code, otl.working_days_in_period(entry.start_date, entry.end_date, member.off_days),
#               entry.daily_hours(member.off_days))
# hours = me.daily_hours(date(2025, 4, 1))
# print(*hours, sep='\n')
# print(sum(h for _, h in hours))
