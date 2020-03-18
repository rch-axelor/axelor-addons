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

import com.axelor.apps.base.db.Address;
import com.axelor.apps.base.db.AppOffice365;
import com.axelor.apps.base.db.ICalendar;
import com.axelor.apps.base.db.ICalendarUser;
import com.axelor.apps.base.db.Partner;
import com.axelor.apps.base.db.PartnerAddress;
import com.axelor.apps.base.db.repo.AddressRepository;
import com.axelor.apps.base.db.repo.AppOffice365Repository;
import com.axelor.apps.base.db.repo.ICalendarRepository;
import com.axelor.apps.base.db.repo.ICalendarUserRepository;
import com.axelor.apps.base.db.repo.PartnerAddressRepository;
import com.axelor.apps.base.db.repo.PartnerRepository;
import com.axelor.apps.base.service.PartnerService;
import com.axelor.apps.crm.db.Event;
import com.axelor.apps.crm.db.EventReminder;
import com.axelor.apps.crm.db.RecurrenceConfiguration;
import com.axelor.apps.crm.db.repo.EventReminderRepository;
import com.axelor.apps.crm.db.repo.EventRepository;
import com.axelor.apps.crm.db.repo.RecurrenceConfigurationRepository;
import com.axelor.apps.crm.service.EventService;
import com.axelor.apps.message.db.EmailAddress;
import com.axelor.apps.message.db.repo.EmailAddressRepository;
import com.axelor.apps.office365.translation.ITranslation;
import com.axelor.auth.AuthUtils;
import com.axelor.auth.db.User;
import com.axelor.auth.db.repo.UserRepository;
import com.axelor.common.StringUtils;
import com.axelor.exception.AxelorException;
import com.axelor.exception.db.repo.TraceBackRepository;
import com.axelor.exception.service.TraceBackService;
import com.axelor.i18n.I18n;
import com.axelor.inject.Beans;
import com.github.scribejava.apis.MicrosoftAzureActiveDirectory20Api;
import com.github.scribejava.core.builder.ServiceBuilder;
import com.github.scribejava.core.model.OAuth2AccessToken;
import com.github.scribejava.core.oauth.OAuth20Service;
import com.google.inject.Inject;
import com.google.inject.persist.Transactional;
import java.net.MalformedURLException;
import java.net.URL;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import wslite.http.HTTPClient;
import wslite.http.HTTPMethod;
import wslite.http.HTTPRequest;
import wslite.http.HTTPResponse;
import wslite.json.JSONArray;
import wslite.json.JSONException;
import wslite.json.JSONObject;

public class Office365ServiceImpl implements Office365Service {

  private static final String CONTACT_URL = "https://graph.microsoft.com/v1.0/me/contacts";
  private static final String CALENDAR_URL = "https://graph.microsoft.com/v1.0/me/calendars";
  private static final String EVENT_URL = "https://graph.microsoft.com/v1.0/me/calendars/%s/events";
  private static final String SCOPE =
      "openid offline_access Contacts.ReadWrite Calendars.ReadWrite";

  @Inject private PartnerService partnerService;

  @Inject private PartnerRepository partnerRepo;
  @Inject private EmailAddressRepository emailAddressRepo;
  @Inject private AddressRepository addressRepo;
  @Inject private PartnerAddressRepository partnerAddressRepo;
  @Inject private EventRepository eventRepo;
  @Inject private ICalendarRepository iCalendarRepo;
  @Inject private ICalendarUserRepository iCalendarUserRepo;
  @Inject private EventReminderRepository eventReminderRepo;
  @Inject private RecurrenceConfigurationRepository recurrenceConfigurationRepo;
  @Inject private UserRepository userRepo;

  public void syncContact(AppOffice365 appOffice365) throws AxelorException, MalformedURLException {

    String accessToken = getAccessTocken(appOffice365);
    URL url = new URL(CONTACT_URL);
    JSONObject jsonObject = fetchData(url, accessToken);
    @SuppressWarnings("unchecked")
    JSONArray jsonArray = (JSONArray) jsonObject.getOrDefault("value", new ArrayList<>());
    for (Object object : jsonArray) {
      jsonObject = (JSONObject) object;
      createContact(jsonObject);
    }
  }

  @SuppressWarnings("unchecked")
  public void syncCalendar(AppOffice365 appOffice365)
      throws AxelorException, MalformedURLException {

    String accessToken = getAccessTocken(appOffice365);
    URL url = new URL(CALENDAR_URL);
    JSONObject jsonObject = fetchData(url, accessToken);
    JSONArray calendarArray = (JSONArray) jsonObject.getOrDefault("value", new ArrayList<>());
    if (calendarArray != null) {
      for (Object object : calendarArray) {
        jsonObject = (JSONObject) object;
        ICalendar iCalendar = createCalendar(jsonObject);
        syncEvent(iCalendar, accessToken);
      }
    }
  }

  @Transactional
  public String getAccessTocken(AppOffice365 appOffice365) throws AxelorException {

    try {
      OAuth20Service authService =
          new ServiceBuilder(appOffice365.getClientId())
              .apiSecret(appOffice365.getClientSecret())
              .callback(appOffice365.getRedirectUri())
              .defaultScope(SCOPE)
              .build(MicrosoftAzureActiveDirectory20Api.instance());
      OAuth2AccessToken accessToken;
      if (StringUtils.isBlank(appOffice365.getRefreshToken())) {
        throw new AxelorException(
            AppOffice365.class,
            TraceBackRepository.CATEGORY_CONFIGURATION_ERROR,
            I18n.get(ITranslation.OFFICE365_TOKEN_ERROR));
      }
      accessToken = authService.refreshAccessToken(appOffice365.getRefreshToken());
      appOffice365.setRefreshToken(accessToken.getRefreshToken());
      Beans.get(AppOffice365Repository.class).save(appOffice365);
      return accessToken.getTokenType() + " " + accessToken.getAccessToken();
    } catch (Exception e) {
      throw new AxelorException(
          AppOffice365.class, TraceBackRepository.CATEGORY_INCONSISTENCY, e.getMessage());
    }
  }

  private JSONObject fetchData(URL url, String accessToken) {

    JSONObject jsonObject = null;
    try {
      HTTPResponse response;
      HTTPClient httpclient = new HTTPClient();
      HTTPRequest request = new HTTPRequest();
      Map<String, Object> headers = new HashMap<>();
      headers.put("Accept", "application/json");
      headers.put("Authorization", accessToken);
      request.setHeaders(headers);
      request.setUrl(url);
      request.setMethod(HTTPMethod.GET);
      response = httpclient.execute(request);
      if (response.getStatusCode() == 200) {
        jsonObject = new JSONObject(response.getContentAsString());
      }
    } catch (Exception e) {
      TraceBackService.trace(e);
    }

    return jsonObject;
  }

  @SuppressWarnings("unchecked")
  private void createContact(JSONObject jsonObject) {

    if (jsonObject != null) {
      try {
        String officeContactId = jsonObject.getOrDefault("id", "").toString();
        Partner partner = partnerRepo.findByOffice365Id(officeContactId);

        if (partner == null) {
          partner = new Partner();
          partner.setOffice365Id(officeContactId);
          partner.setIsContact(true);
          partner.setPartnerTypeSelect(PartnerRepository.PARTNER_TYPE_INDIVIDUAL);
        }
        setPartnerValues(partner, jsonObject);
      } catch (Exception e) {
        TraceBackService.trace(e);
      }
    }
  }

  @SuppressWarnings("unchecked")
  @Transactional
  public void setPartnerValues(Partner partner, JSONObject jsonObject) throws JSONException {

    partner.setFirstName(
        (jsonObject.getOrDefault("givenName", "").toString()
                + " "
                + jsonObject.getOrDefault("middleName", "").toString())
            .trim()
            .replaceAll("null", ""));
    partner.setName(jsonObject.getOrDefault("surname", "").toString().replaceAll("null", ""));
    partner.setFullName(
        jsonObject.getOrDefault("displayName", "").toString().replaceAll("null", ""));
    partner.setCompanyStr(jsonObject.getOrDefault("companyName", "").toString());
    partner.setDepartment(jsonObject.getOrDefault("department", "").toString());
    partner.setMobilePhone(jsonObject.getOrDefault("mobilePhone", "").toString());
    partner.setDescription(jsonObject.getOrDefault("personalNotes", "").toString());
    if (jsonObject.getOrDefault("birthday", null) != null) {
      String birthDateStr = jsonObject.getOrDefault("birthday", "").toString();
      if (!StringUtils.isBlank((birthDateStr)) && !birthDateStr.equals("null")) {
        partner.setBirthdate(LocalDate.parse(birthDateStr.substring(0, birthDateStr.indexOf("T"))));
      }
    }
    partner.setJobTitle(jsonObject.getOrDefault("jobTitle", "").toString());
    partner.setNickName(jsonObject.getOrDefault("nickName", "").toString());

    switch (jsonObject.getOrDefault("title", null).toString().toLowerCase()) {
      case "miss":
        partner.setTitleSelect(PartnerRepository.PARTNER_TITLE_MS);
        break;
      case "dr":
        partner.setTitleSelect(PartnerRepository.PARTNER_TITLE_DR);
        break;
      case "prof":
        partner.setTitleSelect(PartnerRepository.PARTNER_TITLE_PROF);
        break;
      default:
        partner.setTitleSelect(PartnerRepository.PARTNER_TITLE_M);
    }

    JSONArray phones = jsonObject.getJSONArray("homePhones");
    partner.setFixedPhone(
        phones != null && phones.get(0) != null ? phones.get(0).toString() : null);

    manageEmailAddress(jsonObject, partner);
    managePartnerAddress(jsonObject, partner);
    partnerRepo.save(partner);
  }

  @SuppressWarnings("unchecked")
  private void manageEmailAddress(JSONObject jsonObject, Partner partner) {

    try {
      JSONArray emailAddresses = jsonObject.getJSONArray("emailAddresses");
      for (Object object : emailAddresses) {
        JSONObject obj = (JSONObject) object;
        EmailAddress emailAddress =
            emailAddressRepo.findByAddress(obj.getOrDefault("address", "").toString());
        if (emailAddress == null) {
          emailAddress = new EmailAddress();
          emailAddress.setAddress(obj.getOrDefault("address", "").toString());
          emailAddress.setName(obj.getOrDefault("name", "").toString());
          emailAddress.setPartner(partner);
          partner.setEmailAddress(emailAddress);
        } else {
          emailAddress.setName(obj.getOrDefault("name", "").toString());
          emailAddress.setPartner(partner);
          partner.setEmailAddress(emailAddress);
        }
        emailAddressRepo.save(emailAddress);
      }
    } catch (Exception e) {
      TraceBackService.trace(e);
    }
  }

  private void managePartnerAddress(JSONObject jsonObject, Partner partner) {

    try {
      JSONObject homeAddressObj = jsonObject.getJSONObject("homeAddress");
      if (homeAddressObj != null && homeAddressObj.size() > 0) {
        Address defaultAddress = partner != null ? partnerService.getDefaultAddress(partner) : null;
        PartnerAddress partnerAddress = null;
        if (defaultAddress == null) {
          defaultAddress = new Address();
          partnerAddress = new PartnerAddress();
          partnerAddress.setAddress(defaultAddress);
          partnerAddress.setPartner(partner);
          partner.addPartnerAddressListItem(partnerAddress);
        }
        manageAddress(defaultAddress, partner, homeAddressObj);
        if (partnerAddress != null) {
          defaultAddress = partnerAddress.getAddress();
          partnerAddress.setIsDefaultAddr(true);
          partnerAddressRepo.save(partnerAddress);
          partner.setMainAddress(defaultAddress);
        }
      }

      JSONObject businessAddressObj = jsonObject.getJSONObject("businessAddress");
      managePartnerAddress(businessAddressObj, partner, "self.isBusinessAddr = true", true);

      JSONObject otherAddressObj = jsonObject.getJSONObject("otherAddress");
      managePartnerAddress(otherAddressObj, partner, "self.isOtherAddr = true", false);

    } catch (Exception e) {
      TraceBackService.trace(e);
    }
  }

  private void managePartnerAddress(
      JSONObject jsonObject, Partner partner, String filter, boolean isBusinessAddr) {

    PartnerAddress partnerAddress = null;
    if (jsonObject != null && jsonObject.size() > 0) {
      Address address = null;
      if (partner != null) {
        partnerAddress =
            partnerAddressRepo.all().filter(filter + " AND self.partner = ?", partner).fetchOne();
        address = partnerAddress != null ? partnerAddress.getAddress() : null;
      }
      if (address == null) {
        address = new Address();
      }
      if (partnerAddress == null) {
        partnerAddress = new PartnerAddress();
        partnerAddress.setAddress(address);
        partnerAddress.setPartner(partner);
        partner.addPartnerAddressListItem(partnerAddress);
      }
      manageAddress(address, partner, jsonObject);
      if (isBusinessAddr) {
        partnerAddress.setIsBusinessAddr(true);
      } else {
        partnerAddress.setIsOtherAddr(true);
      }
      partnerAddressRepo.save(partnerAddress);
    }
  }

  @SuppressWarnings("unchecked")
  private void manageAddress(Address address, Partner partner, JSONObject jsonAddressObj) {

    try {
      address.setAddressL2(jsonAddressObj.getOrDefault("street", "").toString());
      address.setAddressL3(jsonAddressObj.getOrDefault("city", "").toString());
      address.setAddressL4(jsonAddressObj.getOrDefault("state", "").toString());
      address.setAddressL5(jsonAddressObj.getOrDefault("countryOrRegion", "").toString());
      address.setZip(jsonAddressObj.getOrDefault("postalCode", "").toString());
      addressRepo.save(address);
    } catch (Exception e) {
      TraceBackService.trace(e);
    }
  }

  @SuppressWarnings("unchecked")
  @Transactional
  public ICalendar createCalendar(JSONObject jsonObject) {

    ICalendar iCalendar = null;
    if (jsonObject != null) {
      try {
        iCalendar = iCalendarRepo.findByOffice365Id(jsonObject.getOrDefault("id", "").toString());
        if (iCalendar == null) {
          iCalendar = new ICalendar();
          iCalendar.setOffice365Id(jsonObject.getOrDefault("id", "").toString());
        }
        iCalendar.setName(jsonObject.getOrDefault("name", "").toString());

        JSONObject ownerObject = jsonObject.getJSONObject("owner");
        if (ownerObject != null) {
          String ownerName = ownerObject.getOrDefault("name", "").toString();
          String emailAddressStr = ownerObject.getOrDefault("address", "").toString();
          String code = ownerName.replaceAll("[^a-zA-Z0-9]", "");
          User user = userRepo.findByCode(code);
          if (user == null) {
            user = new User();
            user.setName(ownerName);
            user.setCode(code);
            user.setEmail(emailAddressStr);
            user.setPassword(code);
          }
          iCalendar.setUser(user);
        }
        iCalendarRepo.save(iCalendar);
      } catch (Exception e) {
        TraceBackService.trace(e);
      }
    }

    return iCalendar;
  }

  private void syncEvent(ICalendar iCalendar, String accessToken) throws MalformedURLException {

    if (iCalendar == null || StringUtils.isBlank(iCalendar.getOffice365Id())) {
      return;
    }

    String eventUrl = String.format(EVENT_URL, iCalendar.getOffice365Id());
    URL url = new URL(eventUrl);

    JSONObject jsonObject = fetchData(url, accessToken);
    @SuppressWarnings("unchecked")
    JSONArray eventArray = (JSONArray) jsonObject.getOrDefault("value", null);
    if (eventArray != null) {
      for (Object object : eventArray) {
        jsonObject = (JSONObject) object;
        createEvent(jsonObject, iCalendar);
      }
    }
  }

  @SuppressWarnings("unchecked")
  @Transactional
  public void createEvent(JSONObject jsonObject, ICalendar iCalendar) {

    if (jsonObject != null) {
      try {
        Event event = eventRepo.findByOffice365Id(jsonObject.getOrDefault("id", "").toString());

        if (event == null) {
          event = new Event();
          event.setOffice365Id(jsonObject.getOrDefault("id", "").toString());
          event.setTypeSelect(EventRepository.TYPE_EVENT);
        }

        event.setSubject(jsonObject.getOrDefault("subject", "").toString());
        event.setAllDay((Boolean) jsonObject.getOrDefault("isAllDay", false));

        if ((boolean) jsonObject.getOrDefault("isCancelled", false)) {
          event.setStatusSelect(EventRepository.STATUS_CANCELED);
        } else {
          event.setStatusSelect(EventRepository.STATUS_PLANNED);
        }

        JSONObject startObject = jsonObject.getJSONObject("start");
        event.setStartDateTime(getLocalDateTime(startObject));
        JSONObject endObject = jsonObject.getJSONObject("start");
        event.setEndDateTime(getLocalDateTime(endObject));

        JSONObject bodyObject = jsonObject.getJSONObject("body");
        if (bodyObject != null) {
          event.setDescription(bodyObject.getOrDefault("content", "").toString());
        }

        setEventLocation(event, jsonObject);
        setICalendarUser(event, jsonObject);
        manageReminder(event, jsonObject);
        manageRecurrenceConfigration(event, jsonObject);

        event.setCalendar(iCalendar);
        eventRepo.save(event);
      } catch (Exception e) {
        TraceBackService.trace(e);
      }
    }
  }

  @SuppressWarnings("unchecked")
  private void setEventLocation(Event event, JSONObject jsonObject) throws JSONException {

    String location = "";
    JSONObject locationObject = jsonObject.getJSONObject("location");
    if (locationObject != null && locationObject.containsKey("address")) {
      JSONObject addressObject = locationObject.getJSONObject("address");
      if (addressObject != null) {
        location += addressObject.getOrDefault("street", "").toString();
        location += " " + addressObject.getOrDefault("city", "").toString();
        location += " " + addressObject.getOrDefault("postalCode", "").toString();
        location += " " + addressObject.getOrDefault("state", "").toString();
        location += " " + addressObject.getOrDefault("countryOrRegion", "").toString();
        event.setLocation(location.trim());
      }

      JSONObject coordinateObject = locationObject.getJSONObject("coordinates");
      if (coordinateObject != null && coordinateObject.containsKey("latitude")) {
        String latitude = coordinateObject.getOrDefault("latitude", "").toString();
        String longitude = coordinateObject.getOrDefault("longitude", "").toString();
        event.setGeo(latitude + "," + longitude);
      }
    }
  }

  private void setICalendarUser(Event event, JSONObject jsonObject) throws JSONException {

    JSONArray attendeesArr = jsonObject.getJSONArray("attendees");
    if (attendeesArr != null) {
      for (Object object : attendeesArr) {
        JSONObject attendeeObj = (JSONObject) object;
        @SuppressWarnings("unchecked")
        String type = attendeeObj.getOrDefault("type", "").toString();
        JSONObject emailAddressObj = attendeeObj.getJSONObject("emailAddress");
        ICalendarUser iCalendarUser = getICalendarUser(emailAddressObj, type);
        event.addAttendee(iCalendarUser);
      }
    }

    JSONObject organizerObj = jsonObject.getJSONObject("organizer");
    if (organizerObj != null) {
      JSONObject emailAddressObj = organizerObj.getJSONObject("emailAddress");
      ICalendarUser iCalendarUser = getICalendarUser(emailAddressObj, null);
      event.setOrganizer(iCalendarUser);
    }
  }

  private void manageReminder(Event event, JSONObject jsonObject) {

    @SuppressWarnings("unchecked")
    Integer reminderMinutes = (Integer) jsonObject.getOrDefault("reminderMinutesBeforeStart", null);
    if (reminderMinutes != null) {
      EventReminder eventReminder =
          eventReminderRepo
              .all()
              .filter(
                  "self.modeSelect = ?1 AND self.typeSelect = 1 AND self.durationTypeSelect = ?2 AND self.duration = ?3",
                  EventReminderRepository.MODE_BEFORE_DATE,
                  EventReminderRepository.DURATION_TYPE_MINUTES,
                  reminderMinutes)
              .fetchOne();
      if (eventReminder == null) {
        eventReminder = new EventReminder();
        eventReminder.setModeSelect(EventReminderRepository.MODE_BEFORE_DATE);
        eventReminder.setTypeSelect(1);
        eventReminder.setDuration(reminderMinutes);
        eventReminder.setDurationTypeSelect(EventReminderRepository.DURATION_TYPE_MINUTES);
        eventReminder.setAssignToSelect(EventReminderRepository.ASSIGN_TO_ME);
        eventReminder.setUser(AuthUtils.getUser());
        eventReminderRepo.save(eventReminder);
      }
      if (event.getEventReminderList() == null) {
        event.setEventReminderList(new ArrayList<>());
      }
      if (!event.getEventReminderList().contains(eventReminder)) {
        event.addEventReminderListItem(eventReminder);
      }
    }
  }

  @SuppressWarnings("unchecked")
  private void manageRecurrenceConfigration(Event event, JSONObject jsonObject)
      throws JSONException {

    if (jsonObject.containsKey("recurrence")
        && !jsonObject.get("recurrence").equals(JSONObject.NULL)) {
      JSONObject reminderRecurrenceObj = jsonObject.getJSONObject("recurrence");
      if (reminderRecurrenceObj != null) {

        JSONObject patternObj = reminderRecurrenceObj.getJSONObject("pattern");
        JSONObject rangeObj = reminderRecurrenceObj.getJSONObject("range");

        RecurrenceConfiguration recurrenceConfiguration = new RecurrenceConfiguration();
        recurrenceConfiguration.setEndType(RecurrenceConfigurationRepository.END_TYPE_DATE);

        Integer recurrenceType = 0;
        if (patternObj != null) {
          switch (patternObj.getOrDefault("type", "").toString()) {
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

          if (patternObj.containsKey("dayOfWeek")) {
            JSONArray dayOfWeekArr = patternObj.getJSONArray("dayOfWeek");
            if (dayOfWeekArr != null) {

              recurrenceConfiguration.setMonthRepeatType(
                  RecurrenceConfigurationRepository.REPEAT_TYPE_WEEK);
            } else {
              recurrenceConfiguration.setMonthRepeatType(
                  RecurrenceConfigurationRepository.REPEAT_TYPE_MONTH);
            }
          }
        }

        if (rangeObj != null) {
          LocalDate startDate = LocalDate.parse(rangeObj.getOrDefault("startDate", "").toString());
          LocalDate endDate = LocalDate.parse(rangeObj.getOrDefault("endDate", "").toString());
          recurrenceConfiguration.setStartDate(startDate);
          recurrenceConfiguration.setEndDate(endDate);
        }
        event.setRecurrenceConfiguration(recurrenceConfiguration);
        recurrenceConfiguration.setRecurrenceName(
            Beans.get(EventService.class).computeRecurrenceName(recurrenceConfiguration));
        recurrenceConfigurationRepo.save(recurrenceConfiguration);
      }
    }
  }

  @SuppressWarnings("unchecked")
  private ICalendarUser getICalendarUser(JSONObject emailAddressObj, String type) {

    ICalendarUser iCalendarUser = null;
    if (emailAddressObj != null) {
      String address = emailAddressObj.getOrDefault("address", "").toString();
      iCalendarUser = iCalendarUserRepo.findByEmail(address);

      if (iCalendarUser == null) {
        iCalendarUser = new ICalendarUser();
        iCalendarUser.setEmail(address);
        User user = userRepo.findByEmail(address);
        iCalendarUser.setUser(user);
      }
      iCalendarUser.setName(emailAddressObj.getOrDefault("name", "").toString());

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

  @SuppressWarnings("unchecked")
  private LocalDateTime getLocalDateTime(JSONObject jsonObject) {

    LocalDateTime eventTime = null;
    try {
      if (jsonObject != null) {
        String dateStr = jsonObject.getOrDefault("dateTime", "").toString();
        String timeZone = jsonObject.getOrDefault("timeZone", "").toString();
        if (!StringUtils.isBlank(dateStr) && !StringUtils.isBlank(dateStr)) {
          DateTimeFormatter format = DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss.SSSSSSS");
          eventTime =
              LocalDateTime.parse(dateStr, format).atZone(ZoneId.of(timeZone)).toLocalDateTime();
        }
      }
    } catch (Exception e) {
      TraceBackService.trace(e);
    }
    return eventTime;
  }
}
