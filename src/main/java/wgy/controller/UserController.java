package wgy.controller;

import io.swagger.annotations.Api;
import io.swagger.annotations.ApiImplicitParam;
import io.swagger.annotations.ApiOperation;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import wgy.action.read;
import wgy.entity.User;

import java.util.List;

@Api("用户接口")
@RestController
public class UserController {


    @ApiOperation(value = "通过用户名查询用户信息", notes = "通过用户名查询用户信息", produces = "application/json")
    @ApiImplicitParam(name = "name", value = "用户名", paramType = "query", required = true, dataType = "String")
    @RequestMapping(value = "user/name", method = {RequestMethod.GET, RequestMethod.POST})
    public String getUser(String name) throws Exception {
        List<User> date = new read().getDate();
        System.out.println(date.toString());
        return "成功";
    }


}
