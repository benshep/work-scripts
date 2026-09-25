from datetime import date

import otl
from staff import GroupMember

# Relevant codes for the group
magnet_lab = otl.Code('STGA00038', '01', 'Magnet Test Facility', priority=otl.Priority.BALANCING,
                      project_key='300000607233456', task_key='300000607233465')

mars_underpinning = otl.Code('STGA00206', name='MARS Underpinning', priority=otl.Priority.BALANCING)
clara = otl.Code('STGA00266', name='CLARA Mothball', priority=otl.Priority.AGREED,
                 project_key='300000566261596', task_key='100000870542018')
scu = otl.Code('STGA00273', name='Superconducting Undulator', priority=otl.Priority.AGREED,
               project_key='300000105665086', task_key='100000026583506')
sustainable_accelerators = otl.Code('STGA00298', name='Sustainable Acc (incl CESA and NCF)',
                                    project_key='300000105665723', task_key='100000026584748',
                                    priority=otl.Priority.BALANCING)
xfel_rnd = otl.Code('STGA00241', name='XFEL R&D', priority=otl.Priority.BALANCING,
                    project_key='300000155706667', task_key='300000155706896')
novel_acceleration = otl.Code('STGA00242', name='Novel Acceleration', priority=otl.Priority.BALANCING,
                              project_key='300000156820113', task_key='300000156820132')
ai_ml = otl.Code('STGA00243', name='AI/ML/Data Handling', project_key='300000250508299', task_key='300000250508314')
thin_films = otl.Code('STGA00501', name='Cavity SRF Thin Film Preparation & Charact',
                      project_key='300000105519807', task_key='100000026535097',
                      fusion_name='8Cavity SRF thinfilm prep', priority=otl.Priority.BALANCING)
ruedi_new_code = otl.Code('STGA02008', '01', name='RUEDI - Post Bridging (from July 25)',
                          project_key='300000105559964', task_key='100000026565926',
                          fusion_name='RUEDI 2nd Bridging', priority=otl.Priority.AGREED)
clepto_pocf = otl.Code('STLA00037', '147', name='CLEPTO POCF', fusion_name='Proof of Concept',
                       project_key='300000105544652', task_key='300000322859558',
                       end=date(2026, 4, 30))
# numbers from PoCF Williams EUV Effort.xlsx
# not confirmed by BID yet (Teams message from PHW 2/4)
beuv_pocf = otl.Code('STLA00037', '151', name='POCF2526-13',
                     fusion_name='Proof of Concept Fund', project_key='300000105544652', task_key='300000397271350',
                     end=date(2027, 1, 31))
ukxfel_continuation = otl.Code('STGA00183', '01', name='UK XFEL Design Study - From Oct 25',
                               project_key='300000311968599', task_key='300000311968610',
                               fusion_name='UKXFEL ASTeC Funded', priority=otl.Priority.BALANCING)

# Cristina to Deepa 19/3/26:
# You can continue to book [to EPAC] till December 26 at the current level.
# I will get back to you next week past that date
epac = otl.Code('STKA01103', '06.01', name='EPAC', fusion_name='EPAC Capital Equipment',
                project_key='300000105555889', task_key='100000026560793')

# Calculations for EPITA: 2025-07-27 INFRA_TECH budget_IFAST2 Permanent Magnets_FINAL.xlsx
# Task 01: Performance Review & Accelerator Quadrupole Specification
# Task 02: Modular ZEPTO Product Line Engineering - starts later
epita = otl.Code('STGA02014', task='1', name='EPITA', project_key='300000563034178', task_key='100000865734592',
                 start=date(2026, 5, 1), end=date(2030, 4, 30))

# LEAPS-TECH: see WP1_INFRA-2025-TECH-02_budget_WP_Sources_v4.xlsx
# and INFRA-2025_TECH-02_description_WP_Sources_v2_EuXFEL_EP_clean_v1.docx
# 3 person-months (0.25 FTE) for STFC in Task 1.1
# Milestone 1: conceptual design completed, due M12 (Aug 2027)
# Deliverable 1: magnetic design of HiTSUP, due M18 (Feb 2028)
# Milestone 2: engineering design completed, due M20 (April 2028)
# Book up to M19 (March 2028)
leaps_tech = otl.Code('STGA02015', name='LEAPS-TECH',  # task 01 for labour, 02 for travel
                      start=date(2026, 9, 1), end=date(2028, 3, 31),
                      project_key='300000795470244', task_key='100001237986101')

# European XFEL second fan: travel is task 02. Email DAK 23/9/26
eu_xfel = otl.Code('STGA02017', name='European XFEL', start=date(2026, 9, 1),
                   project_key='300000795587679', task_key='300000795587689')

liora_phase1a = otl.Code('STGA04017', name='LIORA Phase A', fusion_name='AVO System Completion Study',
                         start=date(2026, 6, 1), end=date(2026, 7, 31))
liora_phase1b = otl.Code('STGA04018', name='LIORA Phase B',
                         # start date: unknown at the moment!
                         start=date(2027, 3, 1), end=date(2027, 3, 31))

new_opps = otl.Code('STGA00300', name='New Opportunities', project_key='300000564616944', task_key='300000564693950')

# DAE extension to end Sep 2026
dae = otl.Code('STGA02000', '03', 'ISPF INDIA DAE', end=date(2026, 9, 30),
               project_key='300000105705826', task_key='100000026596483')

members: list[GroupMember] = [
    GroupMember('Ben Shepherd', 207835,
                person_id=100000020410836, assignment_id=300000117877863,
                title='Mr',
                booking_plan=otl.BookingPlan([
                    otl.Entry(epac, 0.03),
                    otl.Entry(epita, 0.0198),
                    otl.Entry(sustainable_accelerators, 0.75),
                    otl.Entry(dae, 0.04),
                    otl.Entry(scu, 0.16),
                ])),
    GroupMember('Alexander Bainbridge',
                200394,
                'alex.bainbridge@stfc.ac.uk', known_as='Alex B',
                person_id=100000020410917, assignment_id=300000117882174,
                booking_plan=otl.BookingPlan([
                    otl.Entry(epac, 0.15),
                    otl.Entry(ruedi_new_code, 0.3),
                    otl.Entry(clara, 0.07),
                    otl.Entry(epita, 0.09),
                    otl.Entry(liora_phase1a, 0.02),
                    # otl.Entry(liora_phase1b, 0.01),
                    # PM solenoid design for Kiril & Oleg's plasma deposition experiment
                    otl.Entry(thin_films, (4 * 5 * otl.hours_per_day) / otl.hours_per_fte,
                              priority=otl.Priority.AGREED),
                    otl.Entry(magnet_lab),
                ])),
    GroupMember('David Dunning',
                204991,
                'david.dunning@stfc.ac.uk', known_as='Dave',
                person_id=100000020417326, assignment_id=300000117978650,
                booking_plan=otl.BookingPlan([
                    otl.Entry(ukxfel_continuation, 0.07),
                    otl.Entry(xfel_rnd, 0.23),
                    otl.Entry(novel_acceleration),
                    otl.Entry(eu_xfel, 0.4),  # wait until contract signed
                    otl.Entry(beuv_pocf, 243 / otl.hours_per_fte),  # 26/27
                ])),
    GroupMember('Neil Thompson',
                206988,
                person_id=100000020415442, assignment_id=300000117997606,
                booking_plan=otl.BookingPlan([
                    otl.Entry(epac, 0.05),
                    otl.Entry(ukxfel_continuation, 0.03),
                    otl.Entry(xfel_rnd),
                    otl.Entry(eu_xfel, 0.2),  # wait until contract signed
                    # otl.Entry(novel_acceleration),
                    otl.Entry(beuv_pocf, 424 / otl.hours_per_fte),  # 26/27
                ])),
    GroupMember('Alexander Hinton',
                201375,
                'alex.hinton@stfc.ac.uk', known_as='Alex H',
                title='Mr',
                person_id=100000020413904, assignment_id=300000117923738,
                booking_plan=otl.BookingPlan([
                    otl.Entry(scu),
                    otl.Entry(clepto_pocf, otl.hours_per_day * 5 / otl.hours_per_fte),
                    otl.Entry(epita, 0.2178),
                    otl.Entry(leaps_tech, 0.25 * 7 / 19),  # for 27/28: 0.25 * 12/19, see comments above
                    otl.Entry(dae, 0.15),
                ])),
    GroupMember('Amelia Pollard',
                205179,
                person_id=100000020414057, assignment_id=300000117928903,
                known_as='Amy',
                booking_plan=otl.BookingPlan([
                    otl.Entry(ai_ml, 0.8),
                    otl.Entry(new_opps, 0.2),
                ])),
    GroupMember('Matthew King',
                207007,
                'matthew.king@stfc.ac.uk', known_as='Matt',
                title='Mr',
                person_id=100000020413933, assignment_id=300000117929802,
                booking_plan=otl.BookingPlan([
                    otl.Entry(clara, 0.5),
                    otl.Entry(ai_ml, 0.35),
                    otl.Entry(new_opps, 0.15),
                ])),
    GroupMember('Nasiq Ziyan',
                207521,
                title='Mr',
                person_id=100000020417760, assignment_id=300000117981407,
                booking_plan=otl.BookingPlan([
                    otl.Entry(clara, 0.35),
                    otl.Entry(new_opps, 0.65),
                ])),
    GroupMember('Thomas Smith',  # starts 2026-08-17
                304560,
                known_as='Tom',
                title='Mr',
                email='thomas.smith2@stfc.ac.uk',
                person_id=300000399437168, assignment_id=300000735307053,
                booking_plan=otl.BookingPlan([
                    otl.Entry(magnet_lab, start_date=date(2026, 8, 17)),
                ]))
]

if __name__ == '__main__':
    # check_total_ftes(members)
    # print(*[person.name for person in members], sep='\t')
    # person.update_off_days()
    # print(*sorted(list(person.off_days)), sep='\n')
    # print(person.daily_bookings(date.today()))
    # run_otl_calculator()
    for member in members:
        member.booking_plan.convex_levelling()
        member.print_workforce_plan()
        # print(entry.code, otl.working_days_in_period(entry.start_date, entry.end_date, member.off_days),
        #       entry.daily_hours(member.off_days))
# hours = me.daily_hours(date(2025, 4, 1))
# print(*hours, sep='\n')
# print(sum(h for _, h in hours))
