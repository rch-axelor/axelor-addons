/*
 * Axelor Business Solutions
 *
 * Copyright (C) 2020 Axelor (<http://axelor.com>).
 *
 * This program is free software: you can redistribute it and/or  modify
 * it under the terms of the GNU Affero General Public License, version 3,
 * as published by the Free Software Foundation.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Affero General Public License for more details.
 *
 * You should have received a copy of the GNU Affero General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
package com.axelor.apps.office365.service;

import com.axelor.apps.base.db.BaseBatch;
import com.axelor.apps.base.db.ICalendar;
import com.axelor.apps.base.db.ICalendarEvent;
import com.axelor.apps.base.db.ICalendarUser;
import com.axelor.apps.base.db.repo.BaseBatchRepository;
import com.axelor.apps.base.db.repo.ICalendarEventRepository;
import com.axelor.apps.base.db.repo.ICalendarRepository;
import com.axelor.apps.base.db.repo.ICalendarUserRepository;
import com.axelor.apps.base.service.app.AppBaseService;
import com.axelor.apps.crm.db.Event;
import com.axelor.apps.crm.db.EventCategory;
import com.axelor.apps.crm.db.EventReminder;
import com.axelor.apps.crm.db.RecurrenceConfiguration;
import com.axelor.apps.crm.db.repo.EventCategoryRepository;
import com.axelor.apps.crm.db.repo.EventReminderRepository;
import com.axelor.apps.crm.db.repo.EventRepository;
import com.axelor.apps.crm.db.repo.RecurrenceConfigurationRepository;
import com.axelor.apps.crm.service.EventService;
import com.axelor.apps.office.db.OfficeAccount;
import com.axelor.apps.office365.translation.ITranslation;
import com.axelor.auth.AuthUtils;
import com.axelor.auth.db.User;
import com.axelor.auth.db.repo.UserRepository;
import com.axelor.common.ObjectUtils;
import com.axelor.db.Query;
import com.axelor.exception.service.TraceBackService;
import com.axelor.i18n.I18n;
import com.axelor.inject.Beans;
import com.google.inject.Inject;
import com.google.inject.persist.Transactional;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZoneOffset;
import java.time.ZonedDateTime;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import org.apache.commons.lang3.StringUtils;
import wslite.json.JSONArray;
import wslite.json.JSONException;
import wslite.json.JSONObject;

public class Office365CalendarService {

  private static final int DEAFULT_SYNC_DURATION = 10;
  private static final int DEAFULT_RECURRENCE_YEARS = 20;

  @Inject private Office365Service office365Service;
  @Inject private EventService eventService;

  @Inject private ICalendarEventRepository eventRepo;
  @Inject private EventCategoryRepository eventCategoryRepo;
  @Inject private ICalendarRepository iCalendarRepo;
  @Inject private ICalendarUserRepository iCalendarUserRepo;
  @Inject private EventReminderRepository eventReminderRepo;
  @Inject private RecurrenceConfigurationRepository recurrenceConfigurationRepo;
  @Inject private UserRepository userRepo;
  @Inject private BaseBatchRepository baseBatchRepo;

  @SuppressWarnings("unchecked")
  @Transactional
  public ICalendar createCalendar(
      JSONObject jsonObject, OfficeAccount officeAccount, LocalDateTime lastSyncOn) {

    if (jsonObject == null) {
      return null;
    }

    ICalendar iCalendar = null;
    try {
      String office365Id = office365Service.processJsonValue("id", jsonObject);
      iCalendar = iCalendarRepo.findByOffice365Id(office365Id);
      if (iCalendar == null) {
        iCalendar = new ICalendar();
        iCalendar.setOffice365Id(office365Id);
        iCalendar.setOfficeAccount(officeAccount);
        iCalendar.setTypeSelect(ICalendarRepository.OFFICE_365);
        iCalendar.setSynchronizationSelect(ICalendarRepository.CRM_SYNCHRO);
        iCalendar.setSynchronizationDuration(DEAFULT_SYNC_DURATION);
      }

      iCalendar.setIsOfficeDefaultCalendar(
          (boolean) jsonObject.getOrDefault("isDefaultCalendar", false));
      iCalendar.setIsOfficeRemovableCalendar(
          (boolean) jsonObject.getOrDefault("isRemovable", false));
      iCalendar.setIsOfficeEditableCalendar((boolean) jsonObject.getOrDefault("canEdit", false));
      iCalendar.setName(office365Service.processJsonValue("name", jsonObject));
      iCalendar.setUser(officeAccount.getOwnerUser());
      iCalendarRepo.save(iCalendar);
      Office365Service.LOG.debug(
          String.format(
              I18n.get(ITranslation.OFFICE365_OBJECT_SYNC_SUCESS),
              "calendar",
              iCalendar.toString()));
    } catch (Exception e) {
      TraceBackService.trace(e);
    }

    return iCalendar;
  }

  @Transactional
  public void createOffice365Calendar(
      ICalendar calendar, OfficeAccount officeAccount, String accessToken) {

    if (calendar.getIsOfficeDefaultCalendar()) {
      return;
    }

    try {
      JSONObject calendarJsonObject = new JSONObject();
      office365Service.putObjValue(calendarJsonObject, "name", calendar.getName());
      office365Service.putUserEmailAddress(calendar.getUser(), calendarJsonObject, "owner");

      String office365Id =
          office365Service.createOffice365Object(
              Office365Service.CALENDAR_URL,
              calendarJsonObject,
              accessToken,
              calendar.getOffice365Id(),
              "calendars",
              "calendar");

      calendar.setOffice365Id(office365Id);
      calendar.setOfficeAccount(officeAccount);
      iCalendarRepo.save(calendar);
    } catch (Exception e) {
      TraceBackService.trace(e);
    }
  }

  @Transactional
  public void removeCalendar(List<Long> calendarIdList, OfficeAccount officeAccount) {

    if (ObjectUtils.isEmpty(calendarIdList)) {
      return;
    }

    try {
      List<ICalendar> calendars =
          iCalendarRepo
              .all()
              .filter(
                  "self.id NOT IN :ids AND self.office365Id IS NOT NULL AND self.officeAccount = :officeAccount AND COALESCE(self.archived, false) = false")
              .bind("ids", calendarIdList)
              .bind("officeAccount", officeAccount)
              .fetch();
      for (ICalendar iCalendar : calendars) {
        List<BaseBatch> batchList =
            baseBatchRepo
                .all()
                .filter(":calendarId IN (SELECT id FROM self.calendarList)")
                .bind("calendarId", iCalendar.getId())
                .fetch();
        for (BaseBatch baseBatch : batchList) {
          baseBatch.removeCalendarListItem(iCalendar);
          baseBatchRepo.save(baseBatch);
        }

        Map<String, Object> bindingMap = new HashMap<>();
        bindingMap.put("calendar", iCalendar);
        List<ICalendarEvent> eventList =
            getEventList(
                iCalendar.getSynchronizationSelect(), "self.calendar = :calendar", bindingMap);
        removeEvent(eventList, iCalendar, null, null);

        iCalendar.setOffice365Id(null);
        iCalendar.setArchived(true);
        iCalendarRepo.remove(iCalendar);
      }

    } catch (Exception e) {
    }
  }

  @Transactional
  public ICalendarEvent createEvent(
      JSONObject jsonObject,
      OfficeAccount officeAccount,
      ICalendar iCalendar,
      LocalDateTime lastSyncOn,
      LocalDateTime now) {

    if (jsonObject == null) {
      return null;
    }

    try {
      String office365Id = office365Service.processJsonValue("id", jsonObject);
      ICalendarEvent event = eventRepo.findByOffice365Id(office365Id);
      LocalDateTime eventStart = getLocalDateTime(jsonObject.getJSONObject("start"));
      LocalDateTime eventEnd = getLocalDateTime(jsonObject.getJSONObject("end"));
      if (iCalendar.getSynchronizationDuration() > 0) {
        LocalDateTime start = now.minusWeeks(iCalendar.getSynchronizationDuration());
        LocalDateTime end = now.plusWeeks(iCalendar.getSynchronizationDuration());
        if (!isDateWithinRange(eventStart, start, end)
            && !isDateWithinRange(eventEnd, start, end)) {
          if (!checkRecurrenceDateRange(start, end, jsonObject)) {
            return event;
          }
        }
      }

      if (event == null) {
        event = new Event();
        event.setOffice365Id(office365Id);
        event.setTypeSelect(EventRepository.TYPE_EVENT);

      } else if (!iCalendar.getKeepRemote()
          && !office365Service.needUpdation(
              jsonObject, lastSyncOn, event.getCreatedOn(), event.getUpdatedOn())) {
        return event;
      }

      setEventValues(
          jsonObject, event, iCalendar, officeAccount.getOwnerUser(), eventStart, eventEnd, now);

      eventRepo.save(event);
      Office365Service.LOG.debug(
          String.format(
              I18n.get(ITranslation.OFFICE365_OBJECT_SYNC_SUCESS), "event", event.toString()));

      return event;
    } catch (Exception e) {
      TraceBackService.trace(e);
      return null;
    }
  }

  @Transactional
  public void createOffice365Event(
      ICalendarEvent event,
      OfficeAccount officeAccount,
      String accessToken,
      User currentUser,
      LocalDateTime now) {

    try {
      ICalendar calendar = event.getCalendar();
      if (calendar.getOffice365Id() == null) {
        return;
      }

      if (calendar.getSynchronizationDuration() > 0) {
        LocalDateTime start = now.minusWeeks(calendar.getSynchronizationDuration());
        LocalDateTime end = now.plusWeeks(calendar.getSynchronizationDuration());
        if (!isDateWithinRange(event.getStartDateTime(), start, end)
            && !isDateWithinRange(event.getEndDateTime(), start, end)) {
          if (!checkRecurrenceDateRange(start, end, event)) {
            return;
          }
        }
      }

      JSONObject eventJsonObject = setOffice365EventValues(event, currentUser);
      String urlStr = String.format(Office365Service.EVENT_URL, calendar.getOffice365Id());
      String office365Id =
          office365Service.createOffice365Object(
              urlStr, eventJsonObject, accessToken, event.getOffice365Id(), "events", "event");
      event.setOffice365Id(office365Id);
      event.setLastOfficeSyncOn(now);
      eventRepo.save(event);

      calendar.setLastSynchronizationDateT(
          Beans.get(AppBaseService.class).getTodayDateTime().toLocalDateTime());
      iCalendarRepo.save(calendar);
    } catch (Exception e) {
      TraceBackService.trace(e);
    }
  }

  @Transactional
  public void removeEvent(
      List<ICalendarEvent> removalEventList,
      ICalendar iCalendar,
      LocalDateTime now,
      List<Long> removedEventIdList) {

    if (ObjectUtils.isEmpty(removalEventList)) {
      return;
    }

    LocalDateTime start = null, end = null;
    if (iCalendar.getSynchronizationDuration() > 0 && now != null) {
      start = now.minusWeeks(iCalendar.getSynchronizationDuration());
      end = now.plusWeeks(iCalendar.getSynchronizationDuration());
    }

    Map<String, Object> bindingMap;
    for (ICalendarEvent event : removalEventList) {
      if (start != null
          && end != null
          && !isDateWithinRange(event.getStartDateTime(), start, end)
          && !isDateWithinRange(event.getEndDateTime(), start, end)) {
        continue;
      }

      if (removedEventIdList != null && !removedEventIdList.contains(event.getId())) {
        removedEventIdList.add(event.getId());
      }
      event.setOffice365Id(null);

      if (iCalendar.getSynchronizationSelect() != null
          && ICalendarRepository.CRM_SYNCHRO.equals(iCalendar.getSynchronizationSelect())) {
        Event crmEvent = (Event) event;
        bindingMap = new HashMap<>();
        bindingMap.put("parent", event);
        List<ICalendarEvent> childEvents =
            getEventList(
                iCalendar.getSynchronizationSelect(), "self.parentEvent = :parent", bindingMap);
        for (ICalendarEvent childEvent : childEvents) {
          if (Event.class.isAssignableFrom(childEvent.getClass())) {
            Event crmChildEvent = (Event) childEvent;
            crmChildEvent.setParentEvent(null);
            Beans.get(EventRepository.class).remove(crmChildEvent);
          }
        }
        Beans.get(EventRepository.class).remove(crmEvent);

      } else {
        eventRepo.remove(event);
      }
    }
  }

  public List<ICalendarEvent> getICalendarEvents(
      ICalendar calendar,
      LocalDateTime lastSync,
      LocalDateTime start,
      List<Long> removedEventIdList) {

    Map<String, Object> bindingMap = new HashMap<>();
    bindingMap.put("calendar", calendar);
    bindingMap.put("start", start);

    // TODO : last fetched event

    String queryStr = "";
    if (lastSync == null) {
      queryStr = "self.createdOn < :start";
    } else {
      queryStr =
          "((self.lastOfficeSyncOn IS NULL OR self.lastOfficeSyncOn < :lastSync OR COALESCE(self.updatedOn, self.createdOn) BETWEEN :lastSync AND :start) "
              + "AND COALESCE(self.calendar.keepRemote, false) = false)";
      bindingMap.put("lastSync", lastSync);
    }
    queryStr =
        "self.calendar = :calendar AND COALESCE(self.archived, false) = false AND ("
            + queryStr
            + " OR self.office365Id IS NULL)";
    if (calendar.getSynchronizationSelect() != null
        && ICalendarRepository.CRM_SYNCHRO.equals(calendar.getSynchronizationSelect())) {
      queryStr += " AND self.parentEvent IS NULL";
    }
    if (ObjectUtils.notEmpty(removedEventIdList)) {
      queryStr += " AND self.id NOT IN :removedIds";
      bindingMap.put("removedIds", removedEventIdList);
    }

    return getEventList(calendar.getSynchronizationSelect(), queryStr, bindingMap);
  }

  public List<ICalendarEvent> getEventList(
      String synchronizationSelect, String filter, Map<String, Object> bindingMap) {

    if (synchronizationSelect != null
        && synchronizationSelect.equals(ICalendarRepository.CRM_SYNCHRO)) {
      Query<Event> eventQuery = Beans.get(EventRepository.class).all().filter(filter);
      for (Entry<String, Object> bindingEntry : bindingMap.entrySet()) {
        eventQuery.bind(bindingEntry.getKey(), bindingEntry.getValue());
      }
      return new ArrayList<ICalendarEvent>(eventQuery.fetch());
    } else {
      Query<ICalendarEvent> eventQuery = eventRepo.all().filter(filter);
      for (Entry<String, Object> bindingEntry : bindingMap.entrySet()) {
        eventQuery.bind(bindingEntry.getKey(), bindingEntry.getValue());
      }
      return eventQuery.fetch();
    }
  }

  private boolean isDateWithinRange(
      LocalDateTime dateTime, LocalDateTime start, LocalDateTime end) {

    if (dateTime == null) {
      return false;
    }

    return !(dateTime.isBefore(start) || dateTime.isAfter(end));
  }

  @SuppressWarnings("unchecked")
  private void setEventValues(
      JSONObject jsonObject,
      ICalendarEvent event,
      ICalendar iCalendar,
      User ownerUser,
      LocalDateTime eventStart,
      LocalDateTime eventEnd,
      LocalDateTime now)
      throws JSONException {

    event.setSubject(jsonObject.getOrDefault("subject", "").toString());
    event.setAllDay((Boolean) jsonObject.getOrDefault("isAllDay", false));
    event.setUser(ownerUser);
    event.setStartDateTime(eventStart);
    event.setEndDateTime(eventEnd);
    event.setCalendar(iCalendar);

    JSONObject bodyObject = jsonObject.getJSONObject("body");
    if (bodyObject != null) {
      event.setDescription(bodyObject.getOrDefault("content", "").toString());
    }

    setVisibilitySelect(jsonObject, event);
    setDisponibilitySelect(jsonObject, event);
    event.setLastOfficeSyncOn(now);

    if (Event.class.isAssignableFrom(event.getClass())) {
      Event crmEvent = (Event) event;
      if ((boolean) jsonObject.getOrDefault("isCancelled", false)) {
        crmEvent.setStatusSelect(EventRepository.STATUS_CANCELED);
      } else {
        crmEvent.setStatusSelect(EventRepository.STATUS_PLANNED);
      }
      setEventLocation(crmEvent, jsonObject);
      setICalendarUser(crmEvent, jsonObject);
      manageReminder(crmEvent, jsonObject);
      manageRecurrenceConfigration(crmEvent, jsonObject);
      manageEventCategory(crmEvent, jsonObject);
    }
  }

  @SuppressWarnings("unchecked")
  private boolean checkRecurrenceDateRange(
      LocalDateTime startDate, LocalDateTime endDate, JSONObject jsonObject) throws JSONException {

    if (!jsonObject.containsKey("recurrence")
        || jsonObject.get("recurrence").equals(JSONObject.NULL)
        || jsonObject.getJSONObject("recurrence") == null) {
      return false;
    }

    JSONObject reminderRecurrenceObj = jsonObject.getJSONObject("recurrence");
    JSONObject patternObj = reminderRecurrenceObj.getJSONObject("pattern");
    JSONObject rangeObj = reminderRecurrenceObj.getJSONObject("range");

    if (patternObj == null || rangeObj == null) {
      return false;
    }

    Integer periodicity = (Integer) patternObj.getOrDefault("interval", 0);
    if (periodicity < 1) {
      return false;
    }

    LocalDate rangeStartDate = LocalDate.parse(rangeObj.getOrDefault("startDate", "").toString());
    if (rangeStartDate == null) {
      return false;
    }

    String endType = (String) rangeObj.getOrDefault("type", null);
    if (StringUtils.isNotBlank(endType) && "endDate".equalsIgnoreCase(endType)) {
      LocalDate rangeEndDate = LocalDate.parse(rangeObj.getOrDefault("endDate", "").toString());
      if (rangeEndDate.isBefore(startDate.toLocalDate())) {
        return false;
      }
    }

    String type = office365Service.processJsonValue("type", patternObj);
    LocalDateTime nextDate = getNextDate(type, rangeStartDate.atStartOfDay(), periodicity);
    while (nextDate.isBefore(endDate)) {
      if (nextDate.isAfter(startDate)) {
        return true;
      }
      nextDate = getNextDate(type, nextDate, periodicity);
    }

    return false;
  }

  private boolean checkRecurrenceDateRange(
      LocalDateTime startDate, LocalDateTime endDate, ICalendarEvent event) {

    if (!Event.class.isAssignableFrom(event.getClass())) {
      return false;
    }

    Event crmEvent = (Event) event;
    RecurrenceConfiguration recurrenceConfiguration = crmEvent.getRecurrenceConfiguration();
    if (recurrenceConfiguration == null || recurrenceConfiguration.getRecurrenceType() == null) {
      return false;
    }

    String type = recurrenceConfiguration.getRecurrenceType().toString();
    Integer periodicity = recurrenceConfiguration.getPeriodicity();
    LocalDateTime nextDate =
        getNextDate(type, recurrenceConfiguration.getStartDate().atStartOfDay(), periodicity);
    while (nextDate.isBefore(endDate)) {
      if (nextDate.isAfter(startDate)) {
        return true;
      }
      nextDate = getNextDate(type, nextDate, periodicity);
    }

    return false;
  }

  private LocalDateTime getNextDate(String type, LocalDateTime startDate, Integer periodicity) {

    switch (type) {
      case "daily":
      case "1":
        return startDate.plusDays(periodicity);
      case "weekly":
      case "2":
        return startDate.plusWeeks(periodicity);
      case "absoluteMonthly":
      case "relativeMonthly":
      case "3":
        return startDate.plusMonths(periodicity);
      case "absoluteYearly":
      case "relativeYearly":
      case "4":
        return startDate.plusYears(periodicity);
      default:
        return startDate;
    }
  }

  private void setVisibilitySelect(JSONObject jsonObject, ICalendarEvent event) {

    String sensitivity = office365Service.processJsonValue("sensitivity", jsonObject);
    Integer visibilitySelect = 0;
    if ("normal".equalsIgnoreCase(sensitivity)) {
      visibilitySelect = ICalendarEventRepository.VISIBILITY_PUBLIC;
    } else if ("private".equalsIgnoreCase(sensitivity)) {
      visibilitySelect = ICalendarEventRepository.VISIBILITY_PRIVATE;
    }
    event.setVisibilitySelect(visibilitySelect);
  }

  private void setDisponibilitySelect(JSONObject jsonObject, ICalendarEvent event) {

    String showAs = office365Service.processJsonValue("showAs", jsonObject);
    Integer disponibilitySelect = 0;
    if ("busy".equalsIgnoreCase(showAs)) {
      disponibilitySelect = ICalendarEventRepository.DISPONIBILITY_BUSY;
    } else if ("free".equalsIgnoreCase(showAs)) {
      disponibilitySelect = ICalendarEventRepository.DISPONIBILITY_AVAILABLE;
    } else if ("oof".equalsIgnoreCase(showAs)) {
      disponibilitySelect = ICalendarEventRepository.DISPONIBILITY_AWAY;
    } else if ("tentative".equalsIgnoreCase(showAs)) {
      disponibilitySelect = ICalendarEventRepository.DISPONIBILITY_TENTATIVE;
    } else if ("workingElsewhere".equalsIgnoreCase(showAs)) {
      disponibilitySelect = ICalendarEventRepository.DISPONIBILITY_WORKING_ELSEWHERE;
    }
    event.setDisponibilitySelect(disponibilitySelect);
  }

  private void setEventLocation(Event event, JSONObject jsonObject) throws JSONException {

    JSONObject locationObject = jsonObject.getJSONObject("location");
    if (locationObject == null) {
      return;
    }

    String location = "";
    if (locationObject.containsKey("address")) {
      JSONObject addressObject = locationObject.getJSONObject("address");
      if (addressObject != null) {
        location += office365Service.processJsonValue("street", addressObject);
        location += " " + office365Service.processJsonValue("city", addressObject);
        location += " " + office365Service.processJsonValue("postalCode", addressObject);
        location += " " + office365Service.processJsonValue("state", addressObject);
        location += " " + office365Service.processJsonValue("countryOrRegion", addressObject);
      }
    } else if (locationObject.containsKey("displayName")) {
      location = office365Service.processJsonValue("displayName", locationObject);
    }
    event.setLocation(location.trim());

    if (locationObject.containsKey("coordinates")) {
      JSONObject coordinateObject = locationObject.getJSONObject("coordinates");
      if (coordinateObject != null && coordinateObject.containsKey("latitude")) {
        String latitude = office365Service.processJsonValue("latitude", coordinateObject);
        String longitude = office365Service.processJsonValue("longitude", coordinateObject);
        event.setGeo(latitude + "," + longitude);
      }
    }
  }

  private void setICalendarUser(Event event, JSONObject jsonObject) throws JSONException {

    JSONArray attendeesArr = jsonObject.getJSONArray("attendees");
    if (attendeesArr != null) {
      event.clearAttendees();
      for (Object object : attendeesArr) {
        JSONObject attendeeObj = (JSONObject) object;
        String type = office365Service.processJsonValue("type", attendeeObj);
        JSONObject emailAddressObj = attendeeObj.getJSONObject("emailAddress");
        ICalendarUser iCalendarUser = getICalendarUser(emailAddressObj, type);
        event.addAttendee(iCalendarUser);
      }
    } else {
      event.clearAttendees();
    }

    JSONObject organizerObj = jsonObject.getJSONObject("organizer");
    if (organizerObj != null) {
      JSONObject emailAddressObj = organizerObj.getJSONObject("emailAddress");
      ICalendarUser iCalendarUser = getICalendarUser(emailAddressObj, null);
      event.setOrganizer(iCalendarUser);
    } else {
      event.setOrganizer(null);
    }
  }

  @Transactional
  public void manageReminder(Event event, JSONObject jsonObject) {

    @SuppressWarnings("unchecked")
    Integer reminderMinutes = (Integer) jsonObject.getOrDefault("reminderMinutesBeforeStart", null);
    if (reminderMinutes == null) {
      return;
    }

    EventReminder eventReminder = null;
    if (event != null) {
      eventReminder =
          eventReminderRepo
              .all()
              .filter(
                  "self.modeSelect = :modeSelect "
                      + "AND self.typeSelect = 1 "
                      + "AND self.durationTypeSelect = :durationTypeSelect "
                      + "AND self.duration = :duration "
                      + "AND self.event = :event")
              .bind("modeSelect", EventReminderRepository.MODE_BEFORE_DATE)
              .bind("durationTypeSelect", EventReminderRepository.DURATION_TYPE_MINUTES)
              .bind("duration", reminderMinutes)
              .bind("event", event)
              .fetchOne();
      if (eventReminder != null) {
        return;
      }
    }

    eventReminder = new EventReminder();
    eventReminder.setModeSelect(EventReminderRepository.MODE_BEFORE_DATE);
    eventReminder.setTypeSelect(1);
    eventReminder.setDuration(reminderMinutes);
    eventReminder.setDurationTypeSelect(EventReminderRepository.DURATION_TYPE_MINUTES);
    eventReminder.setAssignToSelect(EventReminderRepository.ASSIGN_TO_ME);
    eventReminder.setUser(AuthUtils.getUser());
    eventReminderRepo.save(eventReminder);

    if (event.getEventReminderList() == null) {
      event.setEventReminderList(new ArrayList<>());
    }
    if (!event.getEventReminderList().contains(eventReminder)) {
      event.addEventReminderListItem(eventReminder);
    }
  }

  @Transactional
  @SuppressWarnings("unchecked")
  public void manageRecurrenceConfigration(Event event, JSONObject jsonObject)
      throws JSONException {

    if (!jsonObject.containsKey("recurrence")
        || jsonObject.get("recurrence").equals(JSONObject.NULL)
        || jsonObject.getJSONObject("recurrence") == null) {
      return;
    }

    JSONObject reminderRecurrenceObj = jsonObject.getJSONObject("recurrence");
    JSONObject patternObj = reminderRecurrenceObj.getJSONObject("pattern");

    RecurrenceConfiguration recurrenceConfiguration = null;
    if (event.getRecurrenceConfiguration() == null) {
      recurrenceConfiguration = new RecurrenceConfiguration();
    } else {
      recurrenceConfiguration = event.getRecurrenceConfiguration();
    }
    recurrenceConfiguration.setEndType(RecurrenceConfigurationRepository.END_TYPE_DATE);

    Integer recurrenceType = 0;
    if (patternObj != null) {
      switch (office365Service.processJsonValue("type", patternObj)) {
        case "daily":
          recurrenceType = RecurrenceConfigurationRepository.TYPE_DAY;
          break;
        case "weekly":
          recurrenceType = RecurrenceConfigurationRepository.TYPE_WEEK;
          break;
        case "absoluteMonthly":
        case "relativeMonthly":
          recurrenceType = RecurrenceConfigurationRepository.TYPE_MONTH;
          break;
        case "absoluteYearly":
        case "relativeYearly":
          recurrenceType = RecurrenceConfigurationRepository.TYPE_YEAR;
          break;
      }
      recurrenceConfiguration.setRecurrenceType(recurrenceType);

      Integer periodicity = (Integer) patternObj.getOrDefault("interval", null);
      recurrenceConfiguration.setPeriodicity(periodicity);

      if (patternObj.containsKey("daysOfWeek")) {
        JSONArray daysOfWeekArr = patternObj.getJSONArray("daysOfWeek");
        if (daysOfWeekArr != null) {
          for (Object dayOfWeek : daysOfWeekArr) {
            switch (dayOfWeek.toString().toLowerCase()) {
              case "sunday":
                recurrenceConfiguration.setSunday(true);
                break;
              case "monday":
                recurrenceConfiguration.setMonday(true);
                break;
              case "tuesday":
                recurrenceConfiguration.setTuesday(true);
                break;
              case "wednesday":
                recurrenceConfiguration.setWednesday(true);
                break;
              case "thursday":
                recurrenceConfiguration.setThursday(true);
                break;
              case "friday":
                recurrenceConfiguration.setFriday(true);
                break;
              case "saturday":
                recurrenceConfiguration.setSaturday(true);
                break;
            }
          }
          recurrenceConfiguration.setMonthRepeatType(
              RecurrenceConfigurationRepository.REPEAT_TYPE_WEEK);
        } else {
          recurrenceConfiguration.setMonthRepeatType(
              RecurrenceConfigurationRepository.REPEAT_TYPE_MONTH);
        }
      }
    }

    JSONObject rangeObj = reminderRecurrenceObj.getJSONObject("range");
    if (rangeObj != null) {
      LocalDate startDate = LocalDate.parse(rangeObj.getOrDefault("startDate", "").toString());
      recurrenceConfiguration.setStartDate(startDate);

      LocalDate endDate = null;
      String endType = (String) rangeObj.getOrDefault("type", null);
      if (StringUtils.isNotBlank(endType) && "endDate".equalsIgnoreCase(endType)) {
        endDate = LocalDate.parse(rangeObj.getOrDefault("endDate", "").toString());
        recurrenceConfiguration.setEndDate(endDate);
      }

      if (endDate == null) {
        // set default recurrence period if no endDate or 0001-01-01 is specified
        endDate =
            getNextDate("absoluteYearly", startDate.atStartOfDay(), DEAFULT_RECURRENCE_YEARS)
                .toLocalDate();
        recurrenceConfiguration.setEndDate(endDate);
      }
    }

    recurrenceConfiguration.setRecurrenceName(
        Beans.get(EventService.class).computeRecurrenceName(recurrenceConfiguration));
    event.setRecurrenceConfiguration(recurrenceConfiguration);
    manageSubEventChanges(event, recurrenceConfiguration);
    recurrenceConfigurationRepo.save(recurrenceConfiguration);
  }

  private void manageSubEventChanges(Event event, RecurrenceConfiguration recurrenceConfiguration) {

    try {
      if (recurrenceConfiguration.getId() == null) {
        eventService.generateRecurrentEvents(event, recurrenceConfiguration);
        Beans.get(EventRepository.class)
            .all()
            .filter("self.parentEvent =:event")
            .bind("event", event)
            .update("office365Id", event.getOffice365Id());
      } else {
        eventService.applyChangesToAll(event);
      }
    } catch (Exception e) {
      TraceBackService.trace(e);
    }
  }

  @Transactional
  public void manageEventCategory(Event event, JSONObject jsonObject) throws JSONException {

    JSONArray categoryJsonArr = jsonObject.getJSONArray("categories");
    if (categoryJsonArr == null) {
      return;
    }

    EventCategory eventCategory = null;
    for (Object category : categoryJsonArr) {
      String categoryStr = category.toString();
      eventCategory = eventCategoryRepo.findByName(categoryStr);
      if (eventCategory == null) {
        eventCategory = new EventCategory();
        eventCategory.setName(categoryStr);
        eventCategory.setCode(categoryStr);
        eventCategoryRepo.save(eventCategory);
      }
    }
    event.setEventCategory(eventCategory);
  }

  @Transactional
  public ICalendarUser getICalendarUser(JSONObject emailAddressObj, String type) {

    ICalendarUser iCalendarUser = null;
    if (emailAddressObj != null) {
      String address = office365Service.processJsonValue("address", emailAddressObj);
      iCalendarUser = iCalendarUserRepo.findByEmail(address);

      if (iCalendarUser == null) {
        iCalendarUser = new ICalendarUser();
        iCalendarUser.setEmail(address);
        User user = userRepo.findByEmail(address);
        iCalendarUser.setUser(user);
      }
      iCalendarUser.setName(office365Service.processJsonValue("name", emailAddressObj));

      if (!StringUtils.isBlank(type)) {
        switch (type) {
          case "required":
            iCalendarUser.setStatusSelect(ICalendarUserRepository.STATUS_REQUIRED);
            break;
          case "optional":
            iCalendarUser.setStatusSelect(ICalendarUserRepository.STATUS_OPTIONAL);
            break;
        }
      }
      iCalendarUserRepo.save(iCalendarUser);
    }
    return iCalendarUser;
  }

  private LocalDateTime getLocalDateTime(JSONObject jsonObject) {

    LocalDateTime eventTime = null;
    try {
      if (jsonObject != null) {
        String dateStr = office365Service.processJsonValue("dateTime", jsonObject);
        String timeZone = office365Service.processJsonValue("timeZone", jsonObject);
        if (!StringUtils.isBlank(dateStr) && !StringUtils.isBlank(timeZone)) {
          dateStr = StringUtils.endsWithIgnoreCase(dateStr, "Z") ? dateStr : dateStr + "Z";
          if (!ZoneId.systemDefault().toString().equalsIgnoreCase(timeZone)) {
            eventTime = LocalDateTime.ofInstant(Instant.parse(dateStr), ZoneId.systemDefault());
          } else {
            eventTime = LocalDateTime.parse(dateStr);
          }
        }
      }
    } catch (Exception e) {
      TraceBackService.trace(e);
    }
    return eventTime;
  }

  private JSONObject setOffice365EventValues(ICalendarEvent event, User currentUser)
      throws JSONException {

    JSONObject eventJsonObject = new JSONObject();
    office365Service.putObjValue(eventJsonObject, "subject", event.getSubject());
    eventJsonObject.put("isAllDay", event.getAllDay());

    JSONObject bodyJsonObject = new JSONObject();
    bodyJsonObject.put("content", event.getDescription());
    bodyJsonObject.put("contentType", "HTML");
    eventJsonObject.put("body", (Object) bodyJsonObject);

    String utcTimezone = "UTC";
    putDateTime(eventJsonObject, "start", event.getStartDateTime(), utcTimezone);
    putDateTime(eventJsonObject, "end", event.getEndDateTime(), utcTimezone);
    putSensitivity(event, eventJsonObject);
    putShowAs(event, eventJsonObject);
    putLocation(event, eventJsonObject);
    putOrganizer(event, eventJsonObject, currentUser);
    putAttendees(event, eventJsonObject);

    if (Event.class.isAssignableFrom(event.getClass())) {
      Event crmEvent = (Event) event;
      if (EventRepository.STATUS_CANCELED == crmEvent.getStatusSelect()) {
        eventJsonObject.put("isCancelled", true);
      }
      if (crmEvent.getEventCategory() != null) {
        eventJsonObject.put("categories", new String[] {crmEvent.getEventCategory().getName()});
      }
      putReminder(crmEvent, eventJsonObject);
      putRepeat(crmEvent, eventJsonObject, utcTimezone);
    }

    return eventJsonObject;
  }

  private void putSensitivity(ICalendarEvent event, JSONObject eventJsonObject)
      throws JSONException {

    if (event.getVisibilitySelect() == null) {
      return;
    }

    if (ICalendarEventRepository.VISIBILITY_PUBLIC == event.getVisibilitySelect()) {
      eventJsonObject.put("sensitivity", "normal");
    } else if (ICalendarEventRepository.VISIBILITY_PRIVATE == event.getVisibilitySelect()) {
      eventJsonObject.put("sensitivity", "private");
    }
  }

  private void putShowAs(ICalendarEvent event, JSONObject eventJsonObject) throws JSONException {

    if (event.getDisponibilitySelect() == null) {
      return;
    }

    if (ICalendarEventRepository.DISPONIBILITY_BUSY == event.getDisponibilitySelect()) {
      eventJsonObject.put("showAs", "busy");
    } else if (ICalendarEventRepository.DISPONIBILITY_AVAILABLE == event.getDisponibilitySelect()) {
      eventJsonObject.put("showAs", "free");
    } else if (ICalendarEventRepository.DISPONIBILITY_AWAY == event.getDisponibilitySelect()) {
      eventJsonObject.put("showAs", "oof");
    } else if (ICalendarEventRepository.DISPONIBILITY_TENTATIVE == event.getDisponibilitySelect()) {
      eventJsonObject.put("showAs", "tentative");
    } else if (ICalendarEventRepository.DISPONIBILITY_WORKING_ELSEWHERE
        == event.getDisponibilitySelect()) {
      eventJsonObject.put("showAs", "workingElsewhere");
    } else {
      eventJsonObject.put("showAs", "unknown");
    }
  }

  private void putDateTime(
      JSONObject eventJsonObject, String key, LocalDateTime value, String utcTimezone)
      throws JSONException {

    if (value == null) {
      return;
    }

    JSONObject startJsonObject = new JSONObject();
    startJsonObject.put("dateTime", toZone(value, ZoneId.systemDefault(), ZoneOffset.UTC));
    startJsonObject.put("timeZone", utcTimezone);
    eventJsonObject.put(key, (Object) startJsonObject);
  }

  public LocalDateTime toZone(LocalDateTime time, ZoneId fromZone, ZoneId toZone) {

    ZonedDateTime zonedtime = time.atZone(fromZone);
    ZonedDateTime converted = zonedtime.withZoneSameInstant(toZone);
    return converted.toLocalDateTime();
  }

  private void putLocation(ICalendarEvent event, JSONObject eventJsonObject) throws JSONException {

    JSONObject locationJsonObject = new JSONObject();
    locationJsonObject.put("displayName", event.getLocation());

    String geo = event.getGeo();
    if (StringUtils.isNotBlank(geo)) {
      JSONObject coordinatesJsonObject = new JSONObject();
      coordinatesJsonObject.put("latitude", StringUtils.substringBefore(geo, ";"));
      coordinatesJsonObject.put("longitude", StringUtils.substringAfter(geo, ";"));
      locationJsonObject.put("coordinates", (Object) coordinatesJsonObject);
    }

    eventJsonObject.put("location", (Object) locationJsonObject);
  }

  private void putOrganizer(ICalendarEvent event, JSONObject eventJsonObject, User currentUser)
      throws JSONException {

    ICalendarUser calendarUser = event.getOrganizer();
    if (calendarUser == null) {
      return;
    }

    if (calendarUser.getUser() != null && calendarUser.getUser().equals(currentUser)) {
      eventJsonObject.put("isOrganizer", true);
    }

    JSONObject organizerJsonObject = new JSONObject();
    JSONObject emailJsonObject = new JSONObject();
    office365Service.putObjValue(emailJsonObject, "address", calendarUser.getEmail());
    office365Service.putObjValue(emailJsonObject, "name", calendarUser.getName());
    organizerJsonObject.put("emailAddress", (Object) emailJsonObject);
    eventJsonObject.put("organizer", (Object) organizerJsonObject);
  }

  private void putAttendees(ICalendarEvent event, JSONObject eventJsonObject) throws JSONException {

    if (ObjectUtils.isEmpty(event.getAttendees())) {
      return;
    }

    JSONArray attendeesJsonArr = new JSONArray();
    for (ICalendarUser iCalendarUser : event.getAttendees()) {
      JSONObject attendeeJsonObject = new JSONObject();
      JSONObject emailJsonObject = new JSONObject();
      office365Service.putObjValue(emailJsonObject, "address", iCalendarUser.getEmail());
      office365Service.putObjValue(emailJsonObject, "name", iCalendarUser.getName());
      attendeeJsonObject.put("emailAddress", (Object) emailJsonObject);

      if (iCalendarUser.getStatusSelect() == ICalendarUserRepository.STATUS_REQUIRED) {
        attendeeJsonObject.put("type", "required");
      } else {
        attendeeJsonObject.put("type", "optional");
      }
      attendeesJsonArr.add(attendeeJsonObject);
    }
    eventJsonObject.put("attendees", (Object) attendeesJsonArr);
  }

  private void putReminder(Event event, JSONObject eventJsonObject) throws JSONException {

    if (ObjectUtils.isEmpty(event.getEventReminderList())) {
      return;
    }

    EventReminder eventReminder = event.getEventReminderList().get(0);
    Integer duration = 0;
    if (eventReminder.getModeSelect() == EventReminderRepository.MODE_BEFORE_DATE) {
      duration = eventReminder.getDuration();
      if (eventReminder.getDurationTypeSelect() == EventReminderRepository.DURATION_TYPE_HOURS) {
        duration = duration * 60;
      } else if (eventReminder.getDurationTypeSelect()
          == EventReminderRepository.DURATION_TYPE_DAYS) {
        duration = duration * 60 * 24;
      } else if (eventReminder.getDurationTypeSelect()
          == EventReminderRepository.DURATION_TYPE_WEEKS) {
        duration = duration * 60 * 24 * 7;
      }
    } else if (eventReminder.getModeSelect() == EventReminderRepository.MODE_AT_DATE) {
      duration =
          (int)
              eventReminder
                  .getSendingDateT()
                  .until(eventReminder.getEvent().getStartDateTime(), ChronoUnit.MINUTES);
    }
    eventJsonObject.put("reminderMinutesBeforeStart", duration);
    eventJsonObject.put("isReminderOn", true);
  }

  private void putRepeat(Event event, JSONObject eventJsonObject, String timezone)
      throws JSONException {

    if (event.getRecurrenceConfiguration() == null) {
      return;
    }

    RecurrenceConfiguration recurrenceConfg = event.getRecurrenceConfiguration();

    JSONObject rangeJsonObject = new JSONObject();
    LocalDate startOn = recurrenceConfg.getStartDate();
    rangeJsonObject.put("startDate", startOn != null ? startOn.toString() : null);
    rangeJsonObject.put("recurrenceTimeZone", timezone);
    if (recurrenceConfg.getEndType() == RecurrenceConfigurationRepository.END_TYPE_DATE) {
      if (recurrenceConfg.getEndDate() != null) {
        LocalDate endOn = recurrenceConfg.getEndDate();
        rangeJsonObject.put("endDate", endOn != null ? endOn.toString() : null);
        rangeJsonObject.put("type", "endDate");
      } else {
        rangeJsonObject.put("type", "noEnd");
      }
    } else {
      rangeJsonObject.put("numberOfOccurrences", recurrenceConfg.getRepetitionsNumber());
      rangeJsonObject.put("type", "numbered");
    }

    List<String> weeks = new ArrayList<>();
    addWeek(weeks, recurrenceConfg.getSunday(), "sunday");
    addWeek(weeks, recurrenceConfg.getMonday(), "monday");
    addWeek(weeks, recurrenceConfg.getTuesday(), "tuesday");
    addWeek(weeks, recurrenceConfg.getWednesday(), "wednesday");
    addWeek(weeks, recurrenceConfg.getThursday(), "thursday");
    addWeek(weeks, recurrenceConfg.getFriday(), "friday");
    addWeek(weeks, recurrenceConfg.getSaturday(), "saturday");

    JSONObject patternJsonObject = new JSONObject();

    if (recurrenceConfg.getRecurrenceType() == RecurrenceConfigurationRepository.TYPE_DAY) {
      patternJsonObject.put("type", "daily");

    } else if (recurrenceConfg.getRecurrenceType() == RecurrenceConfigurationRepository.TYPE_WEEK) {
      patternJsonObject.put("type", "weekly");
      patternJsonObject.put("daysOfWeek", weeks.toArray());
      patternJsonObject.put("firstDayOfWeek", "sunday");

    } else if (recurrenceConfg.getRecurrenceType()
        == RecurrenceConfigurationRepository.TYPE_MONTH) {
      if (recurrenceConfg.getMonthRepeatType()
          == RecurrenceConfigurationRepository.REPEAT_TYPE_MONTH) {
        patternJsonObject.put("type", "absoluteMonthly");
      } else if (recurrenceConfg.getMonthRepeatType()
          == RecurrenceConfigurationRepository.REPEAT_TYPE_WEEK) {
        patternJsonObject.put("type", "relativeMonthly");
        patternJsonObject.put("daysOfWeek", weeks.toArray());
      }
    } else if (recurrenceConfg.getRecurrenceType() == RecurrenceConfigurationRepository.TYPE_YEAR) {
      patternJsonObject.put("type", "absoluteYearly");
    }
    patternJsonObject.put("interval", recurrenceConfg.getPeriodicity());

    JSONObject recurrenceConfgJsonObject = new JSONObject();
    recurrenceConfgJsonObject.put("pattern", (Object) patternJsonObject);
    recurrenceConfgJsonObject.put("range", (Object) rangeJsonObject);
    eventJsonObject.put("recurrence", (Object) recurrenceConfgJsonObject);
  }

  private void addWeek(List<String> weeks, boolean isDay, String weekDay) {

    if (isDay) {
      weeks.add(weekDay);
    }
  }
}
