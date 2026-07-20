import { describe, expect, it } from "vitest";

import {
  getISOWeek,
  getISOWeekYear,
  getMonday,
} from "../../static/js/filters.js";

// Dates are built from local Y/M/D components; the helpers read those same
// components, so the pinned TZ never shifts the calendar day.
const localDate = (year, month, day) => new Date(year, month - 1, day);

describe("getISOWeek / getISOWeekYear", () => {
  it("numbers a plain mid-year week", () => {
    // 2024-01-01 is a Monday, so it opens ISO week 1 of 2024.
    expect(getISOWeek(localDate(2024, 1, 1))).toBe(1);
    expect(getISOWeekYear(localDate(2024, 1, 1))).toBe(2024);
  });

  it("keeps Sunday in the same week as the preceding Monday", () => {
    // 2024-01-07 (Sun) still belongs to week 1; 2024-01-08 (Mon) starts week 2.
    expect(getISOWeek(localDate(2024, 1, 7))).toBe(1);
    expect(getISOWeek(localDate(2024, 1, 8))).toBe(2);
  });

  it("assigns early-January days to the previous year's last week", () => {
    // 2021-01-01 (Fri): its Thursday is 2020-12-31 -> week 53 of 2020.
    expect(getISOWeek(localDate(2021, 1, 1))).toBe(53);
    expect(getISOWeekYear(localDate(2021, 1, 1))).toBe(2020);
  });

  it("assigns a late-December day to the current year's week 53", () => {
    // 2020-12-31 is itself the Thursday of ISO week 53 of 2020.
    expect(getISOWeek(localDate(2020, 12, 31))).toBe(53);
    expect(getISOWeekYear(localDate(2020, 12, 31))).toBe(2020);
  });

  it("assigns a late-December day to the next year's week 1", () => {
    // 2019-12-30 (Mon): its Thursday is 2020-01-02 -> week 1 of 2020.
    expect(getISOWeek(localDate(2019, 12, 30))).toBe(1);
    expect(getISOWeekYear(localDate(2019, 12, 30))).toBe(2020);
  });

  it("counts week 53 in a year that has one", () => {
    // 2016-01-01 (Fri): its Thursday is 2015-12-31 -> week 53 of 2015.
    expect(getISOWeek(localDate(2016, 1, 1))).toBe(53);
    expect(getISOWeekYear(localDate(2016, 1, 1))).toBe(2015);
  });
});

describe("getMonday", () => {
  const isoParts = (date) => [
    date.getFullYear(),
    date.getMonth() + 1,
    date.getDate(),
  ];

  it("returns the same day when given a Monday", () => {
    expect(isoParts(getMonday(localDate(2024, 3, 4)))).toEqual([2024, 3, 4]);
  });

  it("maps a Sunday back to the Monday of the same week", () => {
    // 2024-03-10 (Sun) -> 2024-03-04 (Mon), exercising the day===0 branch.
    expect(isoParts(getMonday(localDate(2024, 3, 10)))).toEqual([2024, 3, 4]);
  });

  it("crosses a month boundary when the Monday is in the previous month", () => {
    // 2024-03-03 (Sun) -> 2024-02-26 (Mon).
    expect(isoParts(getMonday(localDate(2024, 3, 3)))).toEqual([2024, 2, 26]);
  });
});
