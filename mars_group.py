from datetime import date, timedelta, datetime

import otl
from staff import GroupMember

# Relevant codes for the group
magnet_lab = otl.Code('STGA00029', '02', 'Magnet test facility', priority=otl.Priority.BALANCING)
mars_underpinning = otl.Code('STGA00206', name='MARS Underpinning', priority=otl.Priority.BALANCING)
clara_user_facility = otl.Code('STGA00265', name='CLARA User Facility (incl. machine development) starts Sep 25',
                               fusion_name='CLARA User Facility',
                               priority=otl.Priority.AGREED)
scu = otl.Code('STGA00273', name='Superconducting Undulator', priority=otl.Priority.AGREED)
sustainable_accelerators = otl.Code('STGA00298', name='Sustainable Acc (incl CESA and NCF)',
                                    priority=otl.Priority.BALANCING)
xfel_rnd = otl.Code('STGA00241', name='XFEL R&D', priority=otl.Priority.BALANCING)
novel_acceleration = otl.Code('STGA00242', name='Novel Acceleration', priority=otl.Priority.BALANCING)
ai_ml = xfel_rnd  # Code('STGA09999')  # no code yet
thin_films = otl.Code('STGA00501', name='Cavity SRF Thin Film Preparation & Charact',
                      fusion_name='8Cavity SRF thinfilm prep', priority=otl.Priority.BALANCING)
novel_neg = otl.Code('STGA00502', name='Novel NEG', priority=otl.Priority.BALANCING)
# code changed in November (email from Julian McKenzie 29/10/25)
ruedi_new_code = otl.Code('STGA02011', '02', name='RUEDI new code', fusion_name='RUEDI 2nd Bridging',
                          priority=otl.Priority.AGREED)
clepto_pocf = otl.Code('STLA00037', '147', name='CLEPTO POCF', fusion_name='Proof of Concept',
                       end=date(2026, 4, 30))
# numbers from PoCF Williams EUV Effort.xlsx
beuv_pocf = otl.Code('STLA00037', '151', name='POCF2526-13',
                     fusion_name='Proof of Concept',
                     end=date(2027, 1, 31))
ukxfel_continuation = otl.Code('STGA00183', '01', name='UK XFEL Design Study - From Oct 25',
                               fusion_name='UKXFEL ASTeC',
                               priority=otl.Priority.BALANCING)

# Cristina to Deepa 19/3/26:
# You can continue to book [to EPAC] till December 26 at the current level.
# I will get back to you next week past that date
epac = otl.Code('STKA01103', '06.01', name='EPAC', fusion_name='EPAC Capital')

# Calculations for EPITA: 2025-07-27 INFRA_TECH budget_IFAST2 Permanent Magnets_FINAL.xlsx
epita = otl.Code('no code yet', name='EPITA',
                 start=date(2026, 5, 1), end=date(2030, 4, 30))

# LEAPS-TECH: see WP1_INFRA-2025-TECH-02_budget_WP_Sources_v4.xlsx
# and INFRA-2025_TECH-02_description_WP_Sources_v2_EuXFEL_EP_clean_v1.docx
# 3 person-months (0.25 FTE) for STFC in Task 1.1
# Milestone 1: conceptual design completed, due M12 (Aug 2027)
# Deliverable 1: magnetic design of HiTSUP, due M18 (Feb 2028)
# Milestone 2: engineering design completed, due M20 (April 2028)
# Book up to M19 (March 2028)
leaps_tech = otl.Code('no code yet', name='LEAPS-TECH',
                 start=date(2026, 9, 1), end=date(2028, 3, 31))

members: list[GroupMember] = [
    GroupMember('Ben Shepherd', 207835,
                person_id=100000020410836, assignment_id=300000117877863,
                title='Mr',
                booking_plan=otl.BookingPlan([
                    otl.Entry(epac, 0.05),
                    otl.Entry(ruedi_new_code, 0.0434),
                    otl.Entry(clara_user_facility, 0.25),
                    otl.Entry(epita, 0.0198),
                    otl.Entry(sustainable_accelerators, 0.46),
                    otl.Entry(magnet_lab),
                ])),
    GroupMember('Alexander Bainbridge',
                200394,
                'alex.bainbridge@stfc.ac.uk', known_as='Alex B',
                person_id=100000020410917, assignment_id=300000117882174,
                booking_plan=otl.BookingPlan([
                    otl.Entry(epac, 0.1),
                    otl.Entry(ruedi_new_code, 0.1851),
                    otl.Entry(clara_user_facility, 0.2),
                    otl.Entry(epita, 0.1584),
                    otl.Entry(magnet_lab),
                ])),
    GroupMember('David Dunning',
                204991,
                'david.dunning@stfc.ac.uk', known_as='Dave',
                person_id=100000020417326, assignment_id=300000117978650,
                booking_plan=otl.BookingPlan([
                    otl.Entry(ukxfel_continuation, 0.4),
                    otl.Entry(xfel_rnd, 0.15),
                    otl.Entry(ai_ml),
                    otl.Entry(beuv_pocf, 243 / otl.hours_per_fte),  # 26/27
                ])),
    GroupMember('Neil Thompson',
                206988,
                person_id=100000020415442, assignment_id=300000117997606,
                booking_plan=otl.BookingPlan([
                    otl.Entry(epac, 0.2),
                    otl.Entry(ukxfel_continuation, 0.25),
                    otl.Entry(xfel_rnd, 0.25),
                    otl.Entry(novel_acceleration),
                    otl.Entry(beuv_pocf, 424 / otl.hours_per_fte),  # 26/27
                ])),
    GroupMember('Kiril Marinov',
                204936,
                person_id=100000020410826, assignment_id=300000117884336,
                booking_plan=otl.BookingPlan([
                    otl.Entry(novel_neg, 0.25),
                    otl.Entry(thin_films, 0.5),
                    otl.Entry(mars_underpinning),
                ])),
    GroupMember('Alexander Hinton',
                201375,
                'alex.hinton@stfc.ac.uk', known_as='Alex H',
                title='Mr',
                person_id=100000020413904, assignment_id=300000117923738,
                booking_plan=otl.BookingPlan([
                    otl.Entry(scu),
                    otl.Entry(clepto_pocf, otl.hours_per_day * 5),
                    otl.Entry(epita, 0.2178),
                    otl.Entry(leaps_tech, 0.25 * 7/19)  # for 27/28: 0.25 * 12/19
                ])),
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
                title='Mr',
                person_id=100000020413933, assignment_id=300000117929802,
                booking_plan=otl.BookingPlan([
                    otl.Entry(clara_user_facility, 0.35),
                    otl.Entry(ai_ml),
                ])),
    GroupMember('Nasiq Ziyan',
                207521,
                title='Mr',
                person_id=100000020417760, assignment_id=300000117981407,
                booking_plan=otl.BookingPlan([
                    otl.Entry(clara_user_facility, 0.35),
                    otl.Entry(ai_ml),
                ])),
    GroupMember('Thomas Smith',  # starts 2026-09-07
                0,
                known_as='Tom',
                title='Mr',
                person_id=100000020412335, assignment_id=300000117987914,
                email='thomas.smith@stfc.ac.uk',
                booking_plan=otl.BookingPlan([
                    otl.Entry(magnet_lab, start_date=date(2026, 9, 7)),
                ]))
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
