with open("demofile.txt", "a") as f:
  ZoneNames = ['Slack Bus', 'AK', 'AL','AR', 'AZ', 'CA', 'CO', 'CT', 'DC', 'DE', 'FL', 'GA', 'HI','IA', 'ID', 'IL', 'IN', 'KS', 'KY', 'LA', 'MA', 'ME', 'MD', 'MI', 'MN','MO', 'MS', 'MT', 'NC', 'ND', 'NE', 'NH', 'NJ', 'NM', 'NV', 'NY', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC', 'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WI', 'WV', 'WY']
  ZoneNums= [999, 1, 2, 4, 5, 6, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 44, 45, 46, 47, 48, 49, 50, 51, 53, 54, 55, 56]
  states = dict(zip(ZoneNames, ZoneNums))
  print(states)
  f.write("def StateConverter(state):\n")
  f.write("    stateNum = None\n")
  f.write("    if state == 'Slack Bus':\n")
  f.write("        stateNum = 999")
  for keys in states:
   f.write(f"    elif state == '{keys}':\n")
   f.write(f"        stateNum = {states[keys]}\n")
