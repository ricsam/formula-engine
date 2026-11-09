/**
 * Renders date and time sections of a format code for Excel serial numbers
 * in the 1900 date system: 1 is 1900-01-01, the integer part counts days and
 * the fraction is the time of day. Like Excel, serial 0 is the fictitious
 * 1900-01-00 and serial 60 the fictitious 1900-02-29 (Lotus 1-2-3's leap year
 * bug), so serials from 61 on match real dates.
 */

import type { DateCode, FormatSection } from "./format-code";

const SECONDS_PER_DAY = 86400;

/** The first serial after 9999-12-31, the last date Excel can show. */
const SERIAL_LIMIT = 2958466;

const MONTH_NAMES = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

const WEEKDAY_NAMES = [
  "Sunday",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday",
];

/** The year, month and day of a whole serial number. */
function toYearMonthDay(serial: number): [number, number, number] {
  if (serial === 0) {
    return [1900, 1, 0];
  }
  if (serial === 60) {
    return [1900, 2, 29];
  }
  if (serial < 60) {
    return [1900, serial < 32 ? 1 : 2, ((serial - 1) % 31) + 1];
  }
  // Julian day number to Gregorian date (Fliegel and Van Flandern). Serial 61
  // is 1900-03-01, Julian day 2415080.
  let l = serial + 68569 + 2415019;
  const n = Math.floor((4 * l) / 146097);
  l -= Math.floor((146097 * n + 3) / 4);
  const i = Math.floor((4000 * (l + 1)) / 1461001);
  l = l - Math.floor((1461 * i) / 4) + 31;
  const j = Math.floor((80 * l) / 2447);
  const day = l - Math.floor((2447 * j) / 80);
  l = Math.floor(j / 11);
  const month = j + 2 - 12 * l;
  const year = 100 * (n - 49) + i + l;
  return [year, month, day];
}

const twoDigits = (value: number) => String(value).padStart(2, "0");

/**
 * Renders a date or time section for a serial number, or returns `undefined`
 * when the serial is not a date Excel can show: below 0 or after 9999-12-31.
 *
 * The time is rounded to the precision shown: to whole seconds, or to the
 * fractions of a second of `ss.00`. Minutes and hours are not rounded, so
 * `h:mm` shows 10:30:59 as 10:30, and a time rounded up to midnight moves the
 * date to the next day. A format without time codes shows the date of the
 * serial, unless the time lies within a ten-thousandth of a second of
 * midnight.
 */
export function renderDateSection(
  section: FormatSection,
  serial: number
): string | undefined {
  const layout = section.date!;
  if (serial < 0) {
    return undefined;
  }
  let day = Math.floor(serial);
  const exactSeconds = SECONDS_PER_DAY * (serial - day);
  let seconds = Math.floor(exactSeconds);
  let fraction = exactSeconds - seconds;
  // Floating-point noise around whole seconds
  if (fraction < 1e-6) {
    fraction = 0;
  } else if (fraction > 0.9999) {
    fraction = 0;
    seconds++;
  }
  if (seconds === SECONDS_PER_DAY) {
    seconds = 0;
    day++;
  }
  if (fraction > 0) {
    const decimals = layout.secondDecimals;
    const roundsUp =
      decimals === 0 ? fraction >= 0.5 : fraction > 1 - 0.5 * 10 ** -decimals;
    if (roundsUp) {
      seconds++;
      fraction = 0;
    }
  }
  if (day + seconds / SECONDS_PER_DAY >= SERIAL_LIMIT) {
    return undefined;
  }
  if (seconds === SECONDS_PER_DAY && layout.time) {
    seconds = 0;
    day++;
  }

  const [year, month, dayOfMonth] = toYearMonthDay(day);
  // Excel counts 1900-01-01 as a Sunday (it was a Monday): with the
  // fictitious 1900-02-29, weekdays are right from March 1900 on.
  const weekday = (day + 6) % 7;
  const hour = Math.floor(seconds / 3600);
  const minute = Math.floor(seconds / 60) % 60;
  const second = seconds % 60;
  const subsecondDigits = fraction.toFixed(layout.secondDecimals).slice(2);

  let text = "";
  for (const part of section.parts) {
    switch (part.type) {
      case "date":
        text += renderDateCode(part.code, {
          year,
          month,
          dayOfMonth,
          weekday,
          hour,
          minute,
          second,
          day,
          seconds,
          subsecondDigits,
          twelveHour: layout.twelveHour,
        });
        break;
      case "literal":
        text += part.text;
        break;
      case "point":
        text += ".";
        break;
      default:
        break;
    }
  }
  return text;
}

type DateValues = {
  year: number;
  month: number;
  dayOfMonth: number;
  weekday: number;
  hour: number;
  minute: number;
  second: number;
  /** Whole days, for elapsed time */
  day: number;
  /** Seconds since midnight, for elapsed time */
  seconds: number;
  subsecondDigits: string;
  twelveHour: boolean;
};

function renderDateCode(code: DateCode, values: DateValues): string {
  switch (code.type) {
    case "year":
      return code.digits === 2
        ? twoDigits(values.year % 100)
        : String(values.year).padStart(4, "0");
    case "buddhist-year":
      return code.digits === 2
        ? twoDigits((values.year + 543) % 100)
        : String(values.year + 543);
    case "month":
      return code.pad ? twoDigits(values.month) : String(values.month);
    case "month-name": {
      const name = MONTH_NAMES[values.month - 1]!;
      return code.style === "long"
        ? name
        : code.style === "short"
        ? name.slice(0, 3)
        : name[0]!;
    }
    case "day":
      return code.pad
        ? twoDigits(values.dayOfMonth)
        : String(values.dayOfMonth);
    case "weekday": {
      const name = WEEKDAY_NAMES[values.weekday]!;
      return code.style === "long" ? name : name.slice(0, 3);
    }
    case "hour": {
      const hour = values.twelveHour ? values.hour % 12 || 12 : values.hour;
      return code.pad ? twoDigits(hour) : String(hour);
    }
    case "minute":
      return code.pad ? twoDigits(values.minute) : String(values.minute);
    case "second":
      return code.pad ? twoDigits(values.second) : String(values.second);
    case "subsecond":
      return "." + values.subsecondDigits.slice(0, code.digits);
    case "ampm": {
      const afternoon = values.hour >= 12;
      if (code.letters) {
        return afternoon ? code.letters[1] : code.letters[0];
      }
      return afternoon ? "PM" : "AM";
    }
    case "elapsed": {
      const total =
        code.unit === "hour"
          ? values.day * 24 + Math.floor(values.seconds / 3600)
          : code.unit === "minute"
          ? values.day * 1440 + Math.floor(values.seconds / 60)
          : values.day * SECONDS_PER_DAY + values.seconds;
      return String(total).padStart(code.width, "0");
    }
  }
}
