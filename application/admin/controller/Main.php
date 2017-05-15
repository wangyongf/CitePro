<?php
/*
 * Copyright (C) 1996-2016 YONGF Inc.All Rights Reserved.
 * Scott Wang blog.54yongf.com | blog.csdn.net/yongf2014	
 * 文件名：main.php @ CitePro						
 * 描述：
 * 
 * 修改历史
 * 版本号    作者                     日期                    简要描述
 *  1.0         Scott Wang         16-11-29             新增：Create	
 */

namespace app\admin\controller;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Style\TOC;
use think\Controller;


/**
 * 后台主界面
 *
 * Class main
 *
 * @author      Scott Wang
 * @version     1.0, 16-11-29
 * @since         JSC 1.0
 */
class Main extends Controller
{
    public function main()
    {
        return view();
    }

    /**
     * 登录
     */
    public function signin()
    {
        return view('main/signin');
    }

    /**
     * 注册
     */
    public function signup()
    {
        return view('main/signup');
    }

    /**
     * 生成Word文件
     *
     */
    public function generateWord()
    {
        $phpWord = new PhpWord();

        $section = $phpWord->addSection();

        //预定义样式
        $phpWord->setDefaultFontName("宋体");         //设置默认字体宋体
        $phpWord->setDefaultFontSize(12);               //小四
        $styleTOC = array('tabLeader' => TOC::TAB_LEADER_DOT);
        $phpWord->addTitleStyle(1, array('size' => 20, 'bold' => true, 'name' => '宋体'));
        $phpWord->addTitleStyle(2, array('size' => 16, 'bold' => true, 'name' => '宋体'));
        $phpWord->addTitleStyle(3, array('size' => 12, 'bold' => true, 'name' => '宋体'));
        $phpWord->addTitleStyle(4, array('size' => 8, 'bold' => true, 'name' => '宋体'));
        $phpWord->addTitleStyle(5, array('size' => 4, 'bold' => true, 'name' => '宋体'));

        //添加页脚
        $footer = $section->addFooter();
        $footer->addPreserveText("{PAGE}", array(), array('alignment' => Jc::CENTER));

        //目录
        $section->addText("目录", array(), array('alignment' => Jc::CENTER));
        $section->addTOC(array(), $styleTOC);

//        $section->addPageBreak();

        //添加内容
        $lastOption = -1;
        $count = input('count');                        //项目个数
        for ($i = 1; $i <= $count; $i++) {
            $option = intval(input('select' . $i));
            switch ($option) {
                case 1:
                case 2:
                case 3:
                case 4:
                case 5:
                    if ($option == 1) {
                        $section->addPageBreak();
                    }
                    $section->addTitle(input('input' . $i), $option);
                    $section->addLine();
                    $section->addText(input('textarea' . $i),
                        null,
                        array('indentation' => array('firstLine' => 480))
                    );
                    $section->addLine();

                    $lastOption = $option;
                    break;
                case 6:
                    $section->addLine();
                    $section->addText(input('textarea' . $i),
                        null,
                        array('indentation' => array('firstLine' => 480))
                    );
                    $section->addLine();
//                    $section->addPageBreak();
                    $lastOption = $option;
                    break;
            }
        }

        $fileName = "论文小助手Pro" . date("YmdHis") . ".docx";
        $phpWord->save($fileName, "Word2007", true);
    }
}