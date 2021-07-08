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

import com.axelor.apps.base.db.AppOffice365;
import com.axelor.apps.base.db.ICalendar;
import com.axelor.apps.base.db.ICalendarEvent;
import com.axelor.apps.base.db.Partner;
import com.axelor.apps.base.db.repo.AppOffice365Repository;
import com.axelor.apps.base.db.repo.ICalendarRepository;
import com.axelor.apps.base.service.app.AppBaseService;
import com.axelor.apps.base.service.user.UserService;
import com.axelor.apps.message.db.EmailAddress;
import com.axelor.apps.message.db.Message;
import com.axelor.apps.office.db.OfficeAccount;
import com.axelor.apps.office.db.repo.OfficeAccountRepository;
import com.axelor.apps.office365.translation.ITranslation;
import com.axelor.apps.tool.QueryBuilder;
import com.axelor.auth.db.User;
import com.axelor.auth.db.repo.UserRepository;
import com.axelor.common.ObjectUtils;
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
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import okhttp3.MediaType;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.Request.Builder;
import okhttp3.RequestBody;
import okhttp3.Response;
import org.apache.commons.lang3.StringUtils;
import org.apache.http.client.utils.URIBuilder;
import wslite.http.HTTPClient;
import wslite.http.HTTPMethod;
import wslite.http.HTTPRequest;
import wslite.http.HTTPResponse;
import wslite.json.JSONArray;
import wslite.json.JSONException;
import wslite.json.JSONObject;

public class Office365ServiceImpl implements Office365Service {

  @Inject Office365ContactService contactService;
  @Inject Office365CalendarService calendarService;
  @Inject Office365MailService mailService;

  @Inject AppOffice365Repository appOffice365Repo;
  @Inject OfficeAccountRepository officeAccountRepo;
  @Inject UserRepository userRepo;
  @Inject ICalendarRepository iCalendarRepo;

  private static String query = "(self.office365Id IS NULL AND self.createdOn < :start)";
  private static String lastSyncQuery =
      "(self.office365Id IS NULL OR (COALESCE(self.updatedOn, self.createdOn) BETWEEN :lastSync AND :start))";

  @Override
  @SuppressWarnings("unchecked")
  public String processJsonValue(String key, JSONObject jsonObject) {

    Object value = jsonObject.getOrDefault(key, "");
    if (value == null) {
      return null;
    }
    return value.toString().replaceAll("null", "").trim();
  }

  @Override
  public void putObjValue(JSONObject jsonObject, String key, String value) throws JSONException {

    if (StringUtils.isBlank(value)) {
      return;
    }

    jsonObject.put(key, value);
  }

  @Override
  public LocalDateTime processLocalDateTimeValue(JSONObject jsonObject, String key, ZoneId zoneId) {

    String dateStr = processJsonValue(key, jsonObject);
    if (StringUtils.isBlank(dateStr)) {
      return null;
    }

    LocalDateTime convertedDateTime = LocalDateTime.ofInstant(Instant.parse(dateStr), zoneId);
    return convertedDateTime;
  }

  @Override
  public boolean needUpdation(
      JSONObject jsonObject,
      LocalDateTime lastSyncOn,
      LocalDateTime createdOn,
      LocalDateTime updatedOn) {

    LocalDateTime creationDT =
        processLocalDateTimeValue(jsonObject, "createdDateTime", ZoneId.systemDefault());
    LocalDateTime lastModificationDT =
        processLocalDateTimeValue(jsonObject, "lastModifiedDateTime", ZoneId.systemDefault());
    if ((lastModificationDT != null
            && ((lastSyncOn != null && lastModificationDT.isBefore(lastSyncOn))
                || (updatedOn != null && updatedOn.isAfter(lastModificationDT))
                || (updatedOn == null && createdOn.isAfter(lastModificationDT))))
        || (lastModificationDT == null
            && creationDT != null
            && lastSyncOn != null
            && creationDT.isBefore(lastSyncOn))) {
      return false;
    }

    return true;
  }

  public String createOffice365Object(
      String urlStr,
      JSONObject jsonObject,
      String accessToken,
      String office365Id,
      String key,
      String type) {

    try {
      if (office365Id != null) {
        urlStr = urlStr + "/" + office365Id;
      }
      URL url = new URL(urlStr);
      RequestBody body =
          RequestBody.create(
              jsonObject.toString(), MediaType.parse("application/json; charset=utf-8"));
      Builder builder =
          new Request.Builder()
              .url(url)
              .addHeader("Accept", "application/json")
              .addHeader("Authorization", accessToken);

      if (office365Id != null) {
        builder = builder.patch(body);
      } else {
        builder = builder.post(body);
      }

      Request request = builder.build();
      OkHttpClient httpClient = new OkHttpClient();
      try (Response response = httpClient.newCall(request).execute()) {
        if (response.isSuccessful()) {
          if (office365Id == null) {
            office365Id =
                StringUtils.substringBetween(
                    response.networkResponse().header("Location"), "/" + key + "('", "')");
          }

          jsonObject.put("Office365Id", office365Id);
          Office365Service.LOG.debug(
              String.format(I18n.get(ITranslation.OFFICE365_OBJECT_SYNC_SUCESS), type, jsonObject));
        } else {
          Office365Service.LOG.debug(
              String.format(
                  I18n.get(ITranslation.OFFICE365_OBJECT_SYNC_FAILURE),
                  type,
                  jsonObject,
                  response.code(),
                  response.message(),
                  response.body().string()));
        }
      }
    } catch (Exception e) {
      TraceBackService.trace(e);
    }

    return office365Id;
  }

  public void deleteOffice365Object(
      String urlStr, String office365Id, String accessToken, String type) {

    try {
      if (office365Id != null) {
        urlStr = urlStr + "/" + office365Id;
      }
      URL url = new URL(urlStr);

      Request request =
          new Request.Builder().url(url).addHeader("Authorization", accessToken).delete().build();
      OkHttpClient httpClient = new OkHttpClient();
      try (Response response = httpClient.newCall(request).execute()) {
        if (!response.isSuccessful() && response.code() != 404) {
          TraceBackService.trace(
              new AxelorException(
                  TraceBackRepository.CATEGORY_CONFIGURATION_ERROR,
                  String.format(
                      I18n.get(ITranslation.OFFICE365_OBJECT_REMOVAL_FAILURE),
                      office365Id,
                      response.code(),
                      response.message(),
                      response.body().string())));
        } else {
          Office365Service.LOG.debug(
              String.format(
                  I18n.get(ITranslation.OFFICE365_OBJECT_REMOVAL_SUCESS), type, office365Id));
        }
      }
    } catch (Exception e) {
      TraceBackService.trace(e);
    }
  }

  public void putUserEmailAddress(User user, JSONObject jsonObject, String key)
      throws JSONException {

    if (user == null || (user.getPartner() == null && StringUtils.isNotBlank(user.getEmail()))) {
      return;
    }

    String emailAddressStr = null, emailName = user.getName();
    if (user.getPartner() != null && user.getPartner().getEmailAddress() != null) {
      EmailAddress emailAddress = user.getPartner().getEmailAddress();
      emailAddressStr = emailAddress.getAddress();
      if (StringUtils.isNotBlank(emailAddress.getName())) {
        emailName = emailAddress.getName();
      }
    } else if (StringUtils.isNotBlank(user.getEmail())) {
      emailAddressStr = user.getEmail();
    }

    if (StringUtils.isBlank(emailAddressStr)) {
      return;
    }

    JSONObject emailJsonObj = new JSONObject();
    putObjValue(emailJsonObj, "address", emailAddressStr);
    putObjValue(emailJsonObj, "name", emailName);
    jsonObject.put(key, (Object) emailJsonObj);
  }

  @Transactional
  @Override
  public String getAccessTocken(OfficeAccount officeAccount) throws AxelorException {

    try {
      AppOffice365 appOffice365 = appOffice365Repo.all().fetchOne();
      OAuth20Service authService =
          new ServiceBuilder(appOffice365.getClientId())
              .apiSecret(appOffice365.getClientSecret())
              .callback(appOffice365.getRedirectUri())
              .defaultScope(Office365Service.SCOPE)
              .build(MicrosoftAzureActiveDirectory20Api.instance());
      OAuth2AccessToken accessToken;
      if (StringUtils.isBlank(officeAccount.getRefreshToken())) {
        throw new AxelorException(
            AppOffice365.class,
            TraceBackRepository.CATEGORY_CONFIGURATION_ERROR,
            I18n.get(ITranslation.OFFICE365_TOKEN_ERROR));
      }

      accessToken = authService.refreshAccessToken(officeAccount.getRefreshToken());
      officeAccount.setRefreshToken(accessToken.getRefreshToken());
      appOffice365Repo.save(appOffice365);

      Office365Service.LOG.debug(I18n.get(ITranslation.OFFICE365_ACCESS_TOKEN_SUCESS));
      return accessToken.getTokenType() + " " + accessToken.getAccessToken();
    } catch (Exception e) {
      throw new AxelorException(
          AppOffice365.class, TraceBackRepository.CATEGORY_INCONSISTENCY, e.getMessage());
    }
  }

  @SuppressWarnings("unchecked")
  private JSONArray fetchData(
      String urlStr, String accessToken, boolean isListResult, String type) {

    JSONArray jsonArray = new JSONArray();
    HTTPClient httpclient = new HTTPClient();
    HTTPRequest request = new HTTPRequest();
    Map<String, Object> headers = new HashMap<>();
    headers.put("Accept", "application/json");
    headers.put("Authorization", accessToken);
    request.setHeaders(headers);
    request.setMethod(HTTPMethod.GET);

    Integer top = 500;
    Integer count = 0;
    Integer skip = 0;
    do {
      try {
        URIBuilder ub = new URIBuilder(urlStr);
        ub.addParameter("top", top.toString());
        ub.addParameter("skip", skip.toString());
        ub.addParameter("$count", "true");

        URL url = new URL(ub.toString());
        request.setUrl(url);

        HTTPResponse response = httpclient.execute(request);
        if (response.getStatusCode() == 200) {
          JSONObject jsonObject = new JSONObject(response.getContentAsString());
          if (isListResult) {
            count = (int) jsonObject.getOrDefault("@odata.count", 0);
            jsonArray.addAll((JSONArray) jsonObject.getOrDefault("value", new JSONArray()));
          } else {
            jsonArray.add(jsonObject);
          }
        }
      } catch (Exception e) {
        TraceBackService.trace(e);
      }
      skip += top;
      count -= skip;
    } while (count > 0);

    Office365Service.LOG.debug(
        String.format(
            I18n.get(ITranslation.OFFICE365_OBJECT_FETCH_SUCESS), type, jsonArray.size()));
    return jsonArray;
  }

  @Transactional
  public void syncContact(OfficeAccount officeAccount)
      throws AxelorException, MalformedURLException {

    LocalDateTime start = Beans.get(AppBaseService.class).getTodayDateTime().toLocalDateTime();
    String accessToken = getAccessTocken(officeAccount);

    JSONArray jsonArray = fetchData(Office365Service.CONTACT_URL, accessToken, true, "contacts");
    if (jsonArray != null) {
      List<Long> partnerIdList = new ArrayList<>();
      for (Object object : jsonArray) {
        JSONObject jsonObject = (JSONObject) object;
        Partner partner =
            contactService.createContact(
                jsonObject, officeAccount, officeAccount.getLastContactSyncOn());
        if (partner != null) {
          if (!partnerIdList.contains(partner.getId())) {
            partnerIdList.add(partner.getId());
          }
          if (partner.getMainPartner() != null
              && !partnerIdList.contains(partner.getMainPartner().getId())) {
            partnerIdList.add(partner.getMainPartner().getId());
          }
        }
      }
      contactService.removeContact(partnerIdList, officeAccount);
    }

    String queryStr = query;
    QueryBuilder<Partner> partnerQuery = QueryBuilder.of(Partner.class);
    if (officeAccount.getLastContactSyncOn() != null) {
      queryStr = lastSyncQuery;
      partnerQuery = partnerQuery.bind("lastSync", officeAccount.getLastContactSyncOn());
    }
    List<Partner> partnerList =
        partnerQuery
            .add(queryStr)
            .add(
                "(self.officeAccount = :officeAccount OR self.user.officeAccount = :officeAccount)")
            .bind("start", start)
            .bind("officeAccount", officeAccount)
            .build()
            .fetch();
    if (ObjectUtils.notEmpty(partnerList)) {
      for (Partner partner : partnerList) {
        contactService.createOffice365Contact(partner, officeAccount, accessToken);
      }
    }

    officeAccount.setLastContactSyncOn(
        Beans.get(AppBaseService.class).getTodayDateTime().toLocalDateTime());
    officeAccountRepo.save(officeAccount);
    Office365Service.LOG.debug(I18n.get(ITranslation.OFFICE365_CONTACT_SYNC));
  }

  @Transactional
  public void syncCalendar(OfficeAccount officeAccount)
      throws AxelorException, MalformedURLException {

    LocalDateTime start = Beans.get(AppBaseService.class).getTodayDateTime().toLocalDateTime();
    String accessToken = getAccessTocken(officeAccount);
    List<Long> removedEventIdList = new ArrayList<>();

    JSONArray calendarArray =
        fetchData(Office365Service.CALENDAR_URL, accessToken, true, "calendars");
    if (calendarArray != null) {
      List<Long> calendarIdList = new ArrayList<>();
      for (Object object : calendarArray) {
        JSONObject jsonObject = (JSONObject) object;
        ICalendar iCalendar =
            calendarService.createCalendar(
                jsonObject, officeAccount, officeAccount.getLastCalendarSyncOn());
        if (iCalendar != null && !calendarIdList.contains(iCalendar.getId())) {
          calendarIdList.add(iCalendar.getId());
        }

        syncEvent(
            iCalendar,
            officeAccount,
            accessToken,
            officeAccount.getLastCalendarSyncOn(),
            start,
            removedEventIdList);
        iCalendar.setLastSynchronizationDateT(
            Beans.get(AppBaseService.class).getTodayDateTime().toLocalDateTime());
        iCalendarRepo.save(iCalendar);
      }
      calendarService.removeCalendar(calendarIdList, officeAccount);
    }

    QueryBuilder<ICalendar> calendarQuery =
        QueryBuilder.of(ICalendar.class)
            .add("self.typeSelect = :typeSelect")
            .add(
                "(self.officeAccount = :officeAccount OR self.user.officeAccount = :officeAccount)")
            .add("COALESCE(self.isOfficeDefaultCalendar, false) = false")
            .add("COALESCE(self.archived, false) = false")
            .bind("typeSelect", ICalendarRepository.OFFICE_365)
            .bind("start", start)
            .bind("officeAccount", officeAccount);
    if (officeAccount.getLastCalendarSyncOn() != null) {
      calendarQuery =
          calendarQuery.add(lastSyncQuery).bind("lastSync", officeAccount.getLastCalendarSyncOn());
    } else {
      calendarQuery = calendarQuery.add(query);
    }
    List<ICalendar> calendarList = calendarQuery.build().fetch();
    if (ObjectUtils.notEmpty(calendarList)) {
      for (ICalendar calendar : calendarList) {
        calendarService.createOffice365Calendar(calendar, officeAccount, accessToken);
      }
    }

    List<ICalendar> allCalendar =
        Beans.get(ICalendarRepository.class)
            .all()
            .filter(
                "(self.officeAccount = :officeAccount OR self.user.officeAccount = :officeAccount) "
                    + " AND self.typeSelect = :typeSelect AND self.createdOn < :start "
                    + " AND COALESCE(self.isOfficeEditableCalendar, true) = true")
            .bind("typeSelect", ICalendarRepository.OFFICE_365)
            .bind("officeAccount", officeAccount)
            .bind("start", start)
            .fetch();
    User currentUser = Beans.get(UserService.class).getUser();
    for (ICalendar iCalendar : allCalendar) {
      syncOffice365Event(
          iCalendar, officeAccount, currentUser, accessToken, start, removedEventIdList);
    }

    officeAccount.setLastCalendarSyncOn(
        Beans.get(AppBaseService.class).getTodayDateTime().toLocalDateTime());
    officeAccountRepo.save(officeAccount);
    Office365Service.LOG.debug(I18n.get(ITranslation.OFFICE365_CALENDAR_SYNC));
  }

  @Transactional
  @Override
  public void syncCalendar(ICalendar calendar) throws AxelorException, MalformedURLException {

    OfficeAccount officeAccount = calendar.getOfficeAccount();
    if (officeAccount == null) {
      return;
    }

    LocalDateTime start = Beans.get(AppBaseService.class).getTodayDateTime().toLocalDateTime();
    String accessToken = getAccessTocken(calendar.getOfficeAccount());

    List<Long> removedEventIdList = new ArrayList<>();
    if (StringUtils.isNoneBlank(calendar.getOffice365Id())) {
      JSONArray calendarArray =
          fetchData(
              Office365Service.CALENDAR_URL + "/" + calendar.getOffice365Id(),
              accessToken,
              false,
              "calendars");
      if (calendarArray != null) {
        JSONObject calendarJsonObj = (JSONObject) calendarArray.get(0);
        calendar =
            calendarService.createCalendar(
                calendarJsonObj, officeAccount, officeAccount.getLastCalendarSyncOn());
        syncEvent(
            calendar,
            officeAccount,
            accessToken,
            officeAccount.getLastCalendarSyncOn(),
            start,
            removedEventIdList);
      }
    }

    calendarService.createOffice365Calendar(calendar, officeAccount, accessToken);
    User currentUser = Beans.get(UserService.class).getUser();
    syncOffice365Event(
        calendar, officeAccount, currentUser, accessToken, start, removedEventIdList);

    LocalDateTime lastSync = Beans.get(AppBaseService.class).getTodayDateTime().toLocalDateTime();
    calendar.setLastSynchronizationDateT(lastSync);
    iCalendarRepo.save(calendar);

    officeAccount.setLastCalendarSyncOn(lastSync);
    officeAccountRepo.save(officeAccount);
    Office365Service.LOG.debug(I18n.get(ITranslation.OFFICE365_CALENDAR_SYNC));
  }

  private void syncEvent(
      ICalendar iCalendar,
      OfficeAccount officeAccount,
      String accessToken,
      LocalDateTime lastSyncOn,
      LocalDateTime now,
      List<Long> removedEventIdList)
      throws MalformedURLException {

    if (iCalendar == null || StringUtils.isBlank(iCalendar.getOffice365Id())) {
      return;
    }

    String eventUrl = String.format(Office365Service.EVENT_URL, iCalendar.getOffice365Id());
    JSONArray eventArray = fetchData(eventUrl, accessToken, true, "events");
    if (eventArray != null) {
      List<Long> syncEventIdList = new ArrayList<>();
      for (Object object : eventArray) {
        JSONObject jsonObject = (JSONObject) object;
        ICalendarEvent event =
            calendarService.createEvent(jsonObject, officeAccount, iCalendar, lastSyncOn, now);
        if (event != null && !syncEventIdList.contains(event.getId())) {
          syncEventIdList.add(event.getId());
        }
      }

      String filter = "self.calendar = :calendar AND self.office365Id IS NOT NULL";
      Map<String, Object> bindingMap = new HashMap<>();
      bindingMap.put("calendar", iCalendar);
      if (ObjectUtils.notEmpty(syncEventIdList)) {
        bindingMap.put("ids", syncEventIdList);
        filter += " AND self.id NOT IN :ids";
      }
      if (iCalendar.getSynchronizationSelect() != null
          && ICalendarRepository.CRM_SYNCHRO.equals(iCalendar.getSynchronizationSelect())) {
        filter += " AND self.parentEvent IS NULL";
      }
      List<ICalendarEvent> removalEventList =
          calendarService.getEventList(iCalendar.getSynchronizationSelect(), filter, bindingMap);
      calendarService.removeEvent(removalEventList, iCalendar, now, removedEventIdList);
    }
  }

  private void syncOffice365Event(
      ICalendar calendar,
      OfficeAccount officeAccount,
      User currentUser,
      String accessToken,
      LocalDateTime start,
      List<Long> removedEventIdList) {

    List<ICalendarEvent> eventList =
        calendarService.getICalendarEvents(
            calendar, officeAccount.getLastCalendarSyncOn(), start, removedEventIdList);
    if (ObjectUtils.notEmpty(eventList)) {
      for (ICalendarEvent event : eventList) {
        calendarService.createOffice365Event(event, officeAccount, accessToken, currentUser, start);
      }
    }
  }

  @Transactional
  public void syncMail(OfficeAccount officeAccount, String urlStr)
      throws AxelorException, MalformedURLException {

    LocalDateTime start = Beans.get(AppBaseService.class).getTodayDateTime().toLocalDateTime();
    String accessToken = getAccessTocken(officeAccount);
    JSONArray messageArray = fetchData(urlStr, accessToken, true, "mails");
    if (messageArray != null) {
      List<Long> messageIdList = new ArrayList<>();
      for (Object object : messageArray) {
        JSONObject jsonObject = (JSONObject) object;
        Message message =
            mailService.createMessage(jsonObject, officeAccount, officeAccount.getLastMailSyncOn());
        if (message != null && !messageIdList.contains(message.getId())) {
          messageIdList.add(message.getId());
        }
      }
      mailService.removeMessage(messageIdList, officeAccount);
    }

    String queryStr = query;
    QueryBuilder<Message> messageQuery = QueryBuilder.of(Message.class);
    if (officeAccount.getLastMailSyncOn() != null) {
      queryStr = lastSyncQuery;
      messageQuery = messageQuery.bind("lastSync", officeAccount.getLastMailSyncOn());
    }
    List<Message> messageList =
        messageQuery
            .add(queryStr)
            .add(
                "(self.officeAccount = :officeAccount OR self.senderUser.officeAccount = :officeAccount)")
            .bind("start", start)
            .bind("officeAccount", officeAccount)
            .build()
            .fetch();

    if (ObjectUtils.notEmpty(messageList)) {
      for (Message message : messageList) {
        mailService.createOffice365Mail(message, officeAccount, accessToken);
      }
    }

    officeAccount.setLastMailSyncOn(
        Beans.get(AppBaseService.class).getTodayDateTime().toLocalDateTime());
    officeAccountRepo.save(officeAccount);
    Office365Service.LOG.debug(I18n.get(ITranslation.OFFICE365_MAIL_SYNC));
  }

  public User getUser(String name, String email) {

    String code = name.replaceAll("[^a-zA-Z0-9]", "");
    User user = userRepo.findByCode(code);
    if (user == null) {
      user = new User();
      user.setName(name);
      user.setCode(code);
      user.setEmail(email);
      user.setPassword(code);
    }

    return user;
  }

  @Override
  public void syncUserMail(EmailAddress emailAddress, List<String> emailIds) {

    if (emailAddress == null) {
      return;
    }
    User user = userRepo.findByEmail(emailAddress.getAddress());
    if (user == null) {
      return;
    }
    OfficeAccount officeAccount = user.getOfficeAccount();

    if (ObjectUtils.notEmpty(emailIds)) {
      for (String emailId : emailIds) {
        try {
          this.createUserMail(
              officeAccount,
              String.format(Office365Service.MAIL_ID_URL, emailAddress.getAddress(), emailId));
        } catch (MalformedURLException | AxelorException e) {
          TraceBackService.trace(e);
        }
      }
      return;
    }

    try {

      if (officeAccount != null) {
        Beans.get(Office365Service.class)
            .syncMail(
                officeAccount,
                String.format(Office365Service.MAIL_USER_URL, emailAddress.getAddress()));
      }
    } catch (MalformedURLException | AxelorException e) {
      TraceBackService.trace(e);
    }
  }

  private void createUserMail(OfficeAccount officeAccount, String urlStr)
      throws AxelorException, MalformedURLException {

    if (officeAccount == null) {
      return;
    }

    String accessToken = getAccessTocken(officeAccount);
    JSONArray jsonArr = fetchData(urlStr, accessToken, false, "user");
    if (jsonArr.isEmpty()) {
      return;
    }
    mailService.createMessage((JSONObject) jsonArr.get(0), officeAccount, null);
  }
}
